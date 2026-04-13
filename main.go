package main

import (
	"bytes"
	"context"
	_ "embed"
	"encoding/json"
	"fmt"
	"html"
	"io"
	"log"
	"net"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"regexp"
	"runtime/debug"
	"strconv"
	"strings"
	"sync"
	"syscall"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/getlantern/systray"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/scjalliance/comshim"
	"github.com/shirou/gopsutil/process"
	"golang.org/x/net/proxy"
	"golang.org/x/sys/windows/svc/eventlog"
)

//go:generate winres -output resource.syso version.rc

type Config struct {
	Telegram struct {
		BotToken      string `json:"bot_token"`
		DefaultChatID string `json:"default_chat_id"`
		UseEmojis     bool   `json:"use_emojis"`
	} `json:"telegram"`
	Proxy                ProxyConfig `json:"proxy"`
	CheckIntervalSeconds int         `json:"check_interval_seconds"`
	LoggingEnabled       bool        `json:"logging_enabled"`      // Логирование в приложение
	FileLoggingEnabled   bool        `json:"file_logging_enabled"` // Логирование в файл
	StartMinimized       bool        `json:"start_minimized"`      // Запуск программы в трее
	CutText              string      `json:"cut_text"`
	Folders              []Folder    `json:"folders"`
	IP                   string      `json:"ip"`
	Port                 int         `json:"port"`
}

type ProxyConfig struct {
	Enabled  bool     `json:"enabled"`
	Type     string   `json:"type"` // "http", "https", "socks5"
	Host     string   `json:"host"` // IP или домен
	Port     int      `json:"port"`
	Username string   `json:"username,omitempty"`
	Password string   `json:"password,omitempty"`
	NoProxy  []string `json:"no_proxy,omitempty"`
}

type Folder struct {
	Name          string `json:"name"`
	ChatID        string `json:"chat_id"`
	MessageLength int    `json:"message_length"`
}

// Глобальный клиент (инициализируется при старте)
var httpClient *http.Client

// Инициализация клиента после загрузки конфига
func initHTTPClient(cfg Config) error {
	var err error
	httpClient, err = NewHTTPClientWithProxy(cfg.Proxy)
	return err
}

var (
	config            Config
	processedEmails   = make(map[string]bool)
	mainWindow        fyne.Window        // Основное окно Windows
	mainWindowVisible bool        = true // Флаг для отслеживания видимости главного окна

	showHideMenuItem *systray.MenuItem // Для обновления меню systray
	isWindowVisible  bool              // Переменная для хранения состояния видимости окна

	mutexMsg     sync.Mutex // Мьютекс для безопасного доступа к переменной
	mutexWindows sync.Mutex // Мьютекс для безопасного доступа к переменной

	// Для того, что бы занять порт и программу нельзя было повторно запустить
	listener net.Listener
	// HTTP Сервер для диагностики работы программы
	httpServer *http.Server

	// Для работы с Outlook
	CLSID_OutlookApp = ole.NewGUID("{0006F03A-0000-0000-C000-000000000046}") // GUID класса Outlook.Application
	IID_IDispatch    = ole.IID_IDispatch

	// Для логирования в eventlog
	eventLog *eventlog.Log
)

const httpTimeout = 10 * time.Second

type LogWriter struct {
	entry           *widget.Label
	mutexLog        sync.Mutex
	scrollContainer *container.Scroll
	autoScroll      bool     // Флаг для отслеживания автоматической прокрутки
	maxLines        int      // Максимальное количество строк
	lines           []string // Буфер строк
}

func (lw *LogWriter) Write(p []byte) (n int, err error) {
	lw.mutexLog.Lock()
	defer lw.mutexLog.Unlock()

	newLine := strings.TrimSuffix(string(p), "\n")

	// Добавляем строку в буфер и обрезаем при необходимости
	if len(lw.lines) >= lw.maxLines {
		lw.lines = lw.lines[1:]
	}
	lw.lines = append(lw.lines, newLine)

	// Обновляем текст виджета
	lw.entry.SetText(strings.Join(lw.lines, "\n"))
	lw.entry.Refresh()

	// Автоматическая прокрутка
	if lw.autoScroll && lw.scrollContainer != nil {
		contentHeight := lw.entry.MinSize().Height
		containerHeight := lw.scrollContainer.Size().Height
		offsetY := lw.scrollContainer.Offset.Y

		// Прокручиваем только если пользователь уже внизу
		if contentHeight-offsetY <= containerHeight+10 { // допуск 10 пикселей
			lw.scrollContainer.Offset.Y = contentHeight - containerHeight
			lw.scrollContainer.Refresh()
		}
	}

	return len(p), nil
}

func initEventLog() {
	const sourceName = "OutlookTelegramNotifier"

	// Проверка и регистрация источника событий
	if err := eventlog.InstallAsEventCreate(sourceName, eventlog.Info); err != nil {
		log.Printf("Failed to register event source: %v", err)
	}

	// Открытие логгера
	var err error
	eventLog, err = eventlog.Open(sourceName)
	if err != nil {
		log.Fatalf("Failed to open event log: %v", err)
	}
}

func handleSignals() {
	c := make(chan os.Signal, 1)
	signal.Notify(c, syscall.SIGINT, syscall.SIGTERM, syscall.SIGABRT)

	safeGo(func() {
		sig := <-c
		msg := fmt.Sprintf("Received system signal: %v", sig)
		eventLog.Error(1000, msg)
		os.Exit(1)
	})
}

func handlePanic(r interface{}) {
	// Получаем стек вызовов
	stack := string(debug.Stack())

	// Формируем сообщение об ошибке
	errMsg := fmt.Sprintf("Паника: %v\nСтек вызовов:\n%s", r, stack)

	// Логируем ошибку
	logMessage(errMsg)

	// Записываем ошибку в eventlog
	if eventLog != nil {
		eventLog.Error(1000, errMsg)
	}

	// Записываем ошибку в файл error.log
	logErrorToFile(fmt.Errorf(errMsg))
}

func safeGo(fn func()) {
	go func() {
		defer func() {
			if r := recover(); r != nil {
				// Используем обертку для обработки паник
				handlePanic(r)
			}
		}()
		fn()
	}()
}

func safeGoNoLog(fn func()) {
	go func() {
		defer func() {
			if r := recover(); r != nil {
				// Просто ничего не логируем
			}
		}()
		fn()
	}()
}

func main() {
	// Загрузка конфигурации
	errors := loadConfig("config.json")

	// Формируем адрес для прослушивания
	address := fmt.Sprintf("%s:%d", config.IP, config.Port)

	// Проверяем, запущен ли уже экземпляр приложения
	if isPortInUse(address) {
		// Если приложение уже запущено, показываем окно с ошибкой
		a := app.New()
		w := a.NewWindow("Ошибка")

		// Создаем метку с текстом ошибки
		errorLabel := widget.NewLabel("Приложение уже запущено")

		// Центрируем метку внутри контейнера
		centeredContent := container.NewCenter(errorLabel)

		// Устанавливаем центрированное содержимое в окно
		w.SetContent(centeredContent)

		// Устанавливаем размер окна
		w.Resize(fyne.NewSize(300, 100))

		// Центрируем окно на экране
		w.CenterOnScreen()

		// Показываем окно и запускаем приложение
		w.ShowAndRun()
		return
	}

	// В случае не корректного запуска программы записать ошибку в журнал Windows с кодом 1000
	// Инициализация логгера событий Windows
	initEventLog()
	defer eventLog.Close()

	// Перехват паник
	defer func() {
		if r := recover(); r != nil {
			handlePanic(r)
		}
	}()

	handleSignals()

	// Управления завершением работы программы
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()

	// Инициализация состояния окна
	isWindowVisible = true

	// Создаем приложение
	a := app.New()
	a.Settings().SetTheme(theme.DarkTheme()) // Устанавливаем темную тему

	// Загружаем иконку
	iconName := "assets/icon.png"
	exePath, err := os.Executable()
	if err != nil {
		log.Fatalf("Не удалось получить путь к исполняемому файлу: %v", err)
	}
	iconPath := filepath.Join(filepath.Dir(exePath), iconName)
	iconData, err := os.ReadFile(iconPath)
	iconResource := fyne.NewStaticResource("icon", iconData)

	mainWindow = a.NewWindow("Outlook Telegram Notifier")
	mainWindow.Resize(fyne.NewSize(900, 450)) // Устанавливаем размер окна

	// Устанавливаем иконку для главного окна
	mainWindow.SetIcon(iconResource)

	// Обёртываем окно для отслеживания
	trackedWin := newTrackedWindow(mainWindow)
	mainWindow = trackedWin

	// Создаем текстовое поле для логов
	logMessage("Инициализация текстового поля для логов...")
	logText := widget.NewLabel("")
	logText.Wrapping = fyne.TextWrapWord // Включаем перенос строк по словам
	logText.TextStyle.Monospace = true   // Моноширинный шрифт

	// Контейнер с прокруткой
	scrollContainer := container.NewVScroll(logText)

	// Перенаправляем логи в LogWriter
	logWriter := &LogWriter{
		entry:           logText,
		scrollContainer: scrollContainer,
		autoScroll:      true,                   // Изначально включаем автоматическую прокрутку
		maxLines:        100,                    // Максимальное количество строк
		lines:           make([]string, 0, 100), // Предварительное выделение памяти
	}
	//log.SetOutput(logWriter)

	// Инициализация логгера
	initLogger("otn.log", config.LoggingEnabled, config.FileLoggingEnabled, logWriter)
	defer closeLogger()

	// Отслеживаем изменения позиции скролла
	scrollContainer.OnScrolled = func(offset fyne.Position) {
		contentHeight := logText.MinSize().Height
		containerHeight := scrollContainer.Size().Height

		// Проверяем, находится ли пользователь внизу
		isAtBottom := offset.Y >= contentHeight-containerHeight

		// Если содержимое помещается полностью, всегда включаем автоматическую прокрутку
		if contentHeight <= containerHeight {
			logWriter.autoScroll = true
		} else if isAtBottom {
			// Если пользователь находится внизу, включаем автоматическую прокрутку
			logWriter.autoScroll = true
		} else {
			// Если пользователь прокрутил вверх, отключаем автоматическую прокрутку
			logWriter.autoScroll = false
		}
	}

	// Устанавливаем контейнер с прокруткой как содержимое окна
	mainWindow.SetContent(scrollContainer)

	// Обрабатываем закрытие окна как скрытие
	mainWindow.SetCloseIntercept(func() {
		mainWindow.Hide()
	})

	if len(errors) > 0 {
		// Выводим все ошибки
		for _, err := range errors {
			log.Println(err)
		}
	}

	// Пишем версию программы
	logMessage("Версия программы 1.06")

	// Проверяем значения IP и Port
	// logMessage(fmt.Sprintf("IP из конфигурации: %s", config.IP))
	// logMessage(fmt.Sprintf("Port из конфигурации: %d", config.Port))
	// logMessage(fmt.Sprintf("Формированный адрес для прослушивания: %s", address))

	// Запускаем HTTP-сервер
	safeGo(func() {
		if err := startHTTPServer(address); err != nil {
			logMessage(fmt.Sprintf("Ошибка при запуске HTTP-сервера: %v", err))
		}
	})

	// Инициализация HTTP-клиента с прокси
	if err := initHTTPClient(config); err != nil {
		log.Fatalf("Ошибка инициализации HTTP-клиента: %v", err)
	}

	// Проверка доступа к боту
	if err := checkBotAccess(config.Telegram.BotToken); err != nil {
		logMessage("Нет доступа к Telegram боту: %v", err)
	}

	// Запускаем трей-иконку в отдельной горутине
	// logMessage("Запуск трей-иконки...")
	safeGo(func() {
		systray.Run(onReady, onExit)
	})

	// Сворачивание в трей при запуске, если включено в конфигурации
	if config.StartMinimized {
		logMessage("Сворачиваем программу в трей")
		mainWindow.Hide()
	} else {
		mainWindow.Show()
	}

	// Отслеживаем состояник окна для обновления меню в systray
	// safeGo(monitorWindowState) // закомментировано, что бы часто не обновлялся systray

	// Запуск освновного цикла программы, если нет ошибок в файле конфигурации
	if len(errors) == 0 {
		//startCounter()
		safeGo(func() {
			mainLogic(ctx)
		})
	}

	// Запускаем главный цикл приложения
	a.Run()
}

func startHTTPServer(address string) error {
	if address == "" || strings.HasSuffix(address, ":0") {
		return fmt.Errorf("некорректный адрес для HTTP-сервера: %s", address)
	}

	httpServer = &http.Server{Addr: address}

	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		response := map[string]string{
			"service": "OTN",
			"status":  "UP",
		}
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(response)
	})

	errChan := make(chan error, 1)
	go func() {
		// logMessage(fmt.Sprintf("Запуск HTTP-сервера на адресе: %s", address))
		err := httpServer.ListenAndServe()
		if err != nil && err != http.ErrServerClosed {
			errChan <- fmt.Errorf("Ошибка при запуске HTTP-сервера: %v", err)
		} else {
			errChan <- nil
		}
	}()

	select {
	case err := <-errChan:
		if err != nil {
			logMessage(err.Error())
			return err
		}
	case <-time.After(500 * time.Millisecond):
		logMessage("HTTP-сервер запущен успешно:  %s", address)
	}

	// Ожидание сигнала завершения
	sigChan := make(chan os.Signal, 1)
	signal.Notify(sigChan, os.Interrupt, syscall.SIGTERM)
	<-sigChan

	// Остановка HTTP-сервера
	shutdownHTTPServer()
	return nil
}

func shutdownHTTPServer() {
	if httpServer != nil {
		fmt.Println("Останавливаю HTTP-сервер...")
		ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
		defer cancel()

		if err := httpServer.Shutdown(ctx); err != nil {
			fmt.Printf("Ошибка при остановке HTTP-сервера: %v\n", err)
		} else {
			fmt.Println("HTTP-сервер успешно остановлен")
		}
	}
}

func isPortInUse(address string) bool {
	// Устанавливаем таймаут для подключения
	conn, err := net.DialTimeout("tcp", address, 500*time.Millisecond)
	if err != nil {
		logMessage(fmt.Sprintf("Порт %s свободен: %v", address, err))
		return false
	}
	defer conn.Close()
	logMessage(fmt.Sprintf("Порт %s занят", address))
	return true
}

// func isPortInUse(address string) bool {
// 	// Пытаемся подключиться к указанному адресу
// 	conn, err := net.Dial("tcp", address)
// 	if err != nil {
// 		// Если ошибка, значит порт свободен
// 		return false
// 	}
// 	// Закрываем соединение, если оно было успешно установлено
// 	conn.Close()
// 	// Если удалось подключиться, значит порт занят
// 	return true
// }

func onReady() {
	// Устанавливаем подсказку для иконки в трее
	systray.SetTooltip("Outlook Telegram Notifier")

	// Устанавливаем иконку (необходим файл icon.ico в той же директории)
	iconData := getIcon("assets/icon.ico")
	systray.SetIcon(iconData)

	// Добавляем заголовок и подсказку для иконки
	systray.SetTitle("Outlook Telegram Notifier")

	// Добавляем меню
	showHideMenuItem = systray.AddMenuItem("Показать/скрыть окно", "Показать/скрыть основное окно")
	systray.AddSeparator()
	mQuit := systray.AddMenuItem("Выход", "Завершить программу")

	// Обновляем текст пункта меню при старте
	// updateTrayMenu()

	safeGo(func() {
		for {
			select {
			case <-showHideMenuItem.ClickedCh:
				toggleWindowVisibility()
			case <-mQuit.ClickedCh:
				systray.Quit()
				os.Exit(0)
			}
		}
	})

}

func onExit() {
	// Очищаем ресурсы при выходе
}

func monitorWindowState() {
	ticker := time.NewTicker(1 * time.Second)
	defer ticker.Stop()

	for range ticker.C {
		updateTrayMenu()
	}
}

func updateTrayMenu() {
	// Проверяем фактическое состояние окна
	isVisible := isMainWindowVisible()

	// Обновляем заголовок пункта меню
	if isVisible {
		showHideMenuItem.SetTitle("Скрыть окно")
	} else {
		showHideMenuItem.SetTitle("Показать окно")
	}
}

func toggleWindowVisibility() {

	if mainWindow == nil {
		logMessage("Ошибка: mainWindow не инициализирована")
		return
	}

	// Получаем текущее состояние окна
	isVisible := isMainWindowVisible()

	// Инвертируем состояние видимости окна
	mainWindowVisible = !isVisible

	if mainWindowVisible {
		// Показываем окно
		mainWindow.Show()
		// updateTrayMenu()
	} else {
		// Скрываем окно
		mainWindow.Hide()
		// updateTrayMenu()
	}
}

// Обёртка для отслеживания Show/Hide
type trackedWindow struct {
	fyne.Window
}

func newTrackedWindow(w fyne.Window) *trackedWindow {
	return &trackedWindow{Window: w}
}

func (tw *trackedWindow) Show() {
	mutexWindows.Lock()
	defer mutexWindows.Unlock()

	isWindowVisible = true
	tw.Window.Show()
}

func (tw *trackedWindow) Hide() {
	mutexWindows.Lock()
	defer mutexWindows.Unlock()

	isWindowVisible = false
	tw.Window.Hide()
}

// Функция проверки видимости окна
func isMainWindowVisible() bool {
	mutexWindows.Lock()
	defer mutexWindows.Unlock()

	return isWindowVisible
}

func getIcon(iconName string) []byte {
	// Получение пути к исполняемому файлу
	exePath, err := os.Executable()
	if err != nil {
		logMessage("Не удалось получить путь к исполняемому файлу: %v", err)
		// return getDefaultIcon() // Возвращаем резервную иконку
	}

	// Формирование пути к файлу иконки
	iconPath := filepath.Join(filepath.Dir(exePath), iconName)
	// logMessage("Попытка загрузить иконку из: %s", iconPath)

	// Чтение файла иконки
	iconData, err := os.ReadFile(iconPath)
	if err != nil {
		logMessage("Не удалось прочитать файл иконки: %v", err)
		// return getDefaultIcon() // Возвращаем резервную иконку
	}

	return iconData
}

func loadConfig(filename string) []error {
	var errors []error

	// Получаем путь к исполняемому файлу программы
	exePath, err := os.Executable()
	if err != nil {
		errors = append(errors, fmt.Errorf("Ошибка получения пути к исполняемому файлу: %v", err))
		return errors
	}

	// Получаем директорию, где находится исполняемый файл
	exeDir := filepath.Dir(exePath)

	// Формируем полный путь к файлу конфигурации
	configPath := filepath.Join(exeDir, filename)

	// Чтение файла конфигурации
	data, err := os.ReadFile(configPath)
	if err != nil {
		errors = append(errors, fmt.Errorf("Ошибка чтения конфига: %v", err))
		return errors
	}

	// Парсинг JSON
	if err := json.Unmarshal(data, &config); err != nil {
		errors = append(errors, fmt.Errorf("Ошибка парсинга конфигурации: %v", err))
		return errors
	}

	// Проверка конфигурации
	if err := validateConfig(); err != nil {
		errors = append(errors, fmt.Errorf("Ошибка валидации конфигурации: %v", err))
		return errors
	}

	return errors
}

func validateConfig() error {
	// Проверка BotToken
	if !isValidBotToken(config.Telegram.BotToken) {
		return fmt.Errorf("Некорректный BotToken")
	}

	// Проверка DefaultChatID
	if !isValidChatID(config.Telegram.DefaultChatID) {
		return fmt.Errorf("Некорректный DefaultChatID")
	}

	// Проверка UseEmojis
	if config.Telegram.UseEmojis != true && config.Telegram.UseEmojis != false {
		return fmt.Errorf("UseEmojis должно быть true или false")
	}

	// Проверка CheckIntervalSeconds
	if config.CheckIntervalSeconds < 0 || config.CheckIntervalSeconds > 1000 {
		return fmt.Errorf("CheckIntervalSeconds должно быть в диапазоне от 0 до 1000")
	}

	// Проверка LoggingEnabled
	if config.LoggingEnabled != true && config.LoggingEnabled != false {
		return fmt.Errorf("LoggingEnabled должно быть true или false")
	}

	// Проверка FileLoggingEnabled
	if config.FileLoggingEnabled != true && config.FileLoggingEnabled != false {
		return fmt.Errorf("FileLoggingEnabled должно быть true или false")
	}

	// Проверка StartMinimized
	if config.StartMinimized != true && config.StartMinimized != false {
		return fmt.Errorf("StartMinimized должно быть true или false")
	}

	// Проверка CutText
	if len(config.CutText) != 0 && len(config.CutText) < 4 {
		return fmt.Errorf("CutText должен быть либо пустым, либо иметь длину не менее 4 символов")
	}

	// Проверка Folders
	for i, folder := range config.Folders {
		// Проверка Name
		if len(folder.Name) == 0 || len(folder.Name) > 150 {
			return fmt.Errorf("Name в папке %d должен быть текстом длиной от 1 до 150 символов", i)
		}

		// Проверка ChatID
		if !isValidChatID(folder.ChatID) {
			return fmt.Errorf("Некорректный ChatID в папке %d", i)
		}

		// Проверка MessageLength
		if folder.MessageLength < 0 || folder.MessageLength > 4000 {
			return fmt.Errorf("MessageLength в папке %d должно быть в диапазоне от 0 до 4000", i)
		}
	}

	// Проверка IP
	if config.IP != "127.0.0.1" && config.IP != "0.0.0.0" {
		return fmt.Errorf("IP должен быть либо 127.0.0.1, либо 0.0.0.0")
	}

	// Проверка Port
	if config.Port < 1024 || config.Port > 49151 {
		return fmt.Errorf("Port должен быть в диапазоне от 1024 до 49151")
	}

	// Проверка Proxy (если включен)
	if config.Proxy.Enabled {
		// Тип прокси
		validTypes := map[string]bool{"http": true, "https": true, "socks5": true}
		if !validTypes[config.Proxy.Type] {
			return fmt.Errorf("неподдерживаемый тип прокси: %s (допустимы: http, https, socks5)", config.Proxy.Type)
		}

		// Host
		if config.Proxy.Host == "" {
			return fmt.Errorf("proxy.host не может быть пустым при включенном прокси")
		}

		// Port
		if config.Proxy.Port < 1 || config.Proxy.Port > 65535 {
			return fmt.Errorf("proxy.port должен быть в диапазоне 1-65535, получено: %d", config.Proxy.Port)
		}

		// Если есть пароль, должен быть и логин (опционально, но логично)
		if config.Proxy.Password != "" && config.Proxy.Username == "" {
			return fmt.Errorf("proxy: пароль указан, но логин пустой")
		}
	}

	return nil
}

func isValidBotToken(token string) bool {
	// Простая проверка формата BotToken (пример: "123456789:ABCdefGhIJKlmNoPQRstuVWXyz")
	re := regexp.MustCompile(`^\d+:[\w-]+$`)
	return re.MatchString(token)
}

func isValidChatID(chatID string) bool {
	// Проверяем, является ли chatID пустой строкой или "0"
	if chatID == "" || chatID == "0" {
		return true
	}

	// ChatID может быть числом (например, "-1001234567890") или строкой (например, "@channel_name")
	if _, err := strconv.ParseInt(chatID, 10, 64); err == nil {
		return true
	}
	return regexp.MustCompile(`^@[a-zA-Z0-9_]+$`).MatchString(chatID)
}

func checkBotAccess(botToken string) error {
	// Проверка доступа к боту через API Telegram
	url := fmt.Sprintf("https://api.telegram.org/bot%s/getMe", botToken)
	resp, err := httpClient.Get(url) // ← Используем клиент с прокси
	if err != nil {
		return fmt.Errorf("ошибка HTTP-запроса: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("статус ответа: %s", resp.Status)
	}

	var result map[string]interface{}
	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return fmt.Errorf("ошибка декодирования ответа: %v", err)
	}

	ok, exists := result["ok"].(bool)
	if !exists || !ok {
		return fmt.Errorf("бот недоступен")
	}

	return nil
}

var logFile *os.File

type MultiWriter struct {
	writers []io.Writer
}

func (mw *MultiWriter) Write(p []byte) (n int, err error) {
	for _, w := range mw.writers {
		if w != nil {
			if _, err := w.Write(p); err != nil {
				return 0, err
			}
			if flusher, ok := w.(interface{ Flush() error }); ok {
				flusher.Flush()
			}
		}
	}
	return len(p), nil
}

func initLogger(logFilePath string, consoleLoggingEnabled bool, fileLoggingEnabled bool, logWriter *LogWriter) {
	var writers []io.Writer

	// Логирование в файл
	if fileLoggingEnabled {
		file, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
		if err != nil {
			log.Fatalf("Не удалось открыть файл логов: %v", err)
		}
		writers = append(writers, file)
	}

	// Добавляем LogWriter для GUI, если включено логирование
	if consoleLoggingEnabled {
		if logWriter != nil {
			writers = append(writers, logWriter)
		}
	}

	// Создаем мультиплексированный writer
	multiWriter := &MultiWriter{writers: writers}

	// Устанавливаем новый вывод для логов
	log.SetOutput(multiWriter)
}

func (mw *MultiWriter) Close() {
	for _, w := range mw.writers {
		if closer, ok := w.(io.Closer); ok {
			closer.Close()
		}
	}
}

func closeLogger() {
	// Получаем текущий вывод логов
	if multiWriter, ok := log.Writer().(*MultiWriter); ok {
		multiWriter.Close()
	}
}

func logMessage(format string, args ...interface{}) {
	// Форматируем сообщение
	message := fmt.Sprintf(format, args...)

	// Логгируем с использованием настроек из initLogger
	log.Println(message)
}

func logErrorToFile(err error) {
	// Открываем файл error.log в режиме добавления
	file, openErr := os.OpenFile("error.log", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if openErr != nil {
		logMessage("Не удалось открыть файл error.log: %v", openErr)
		return
	}
	defer file.Close()

	// Записываем ошибку в файл
	timestamp := time.Now().Format("2006-01-02 15:04:05")
	logMSG := fmt.Sprintf("[%s] %v\n", timestamp, err)
	if _, writeErr := file.WriteString(logMSG); writeErr != nil {
		logMessage("Не удалось записать в файл error.log: %v", writeErr)
	}
}

var semaphore = make(chan struct{}, 1) // Ограничение до 1 одновременных горутин

// Основная логика программы и ее функции
func mainLogic(ctx context.Context) {
	logMessage("Приложение успешно запущено и готово к работе")
	eventLog.Info(0, "Приложение успешно запущено и готово к работе")

	// Гибкое управление COM потоками || Глобальная инициализация COM
	comshim.Add(1)
	defer comshim.Done()

	// Инициализация COM с обработкой ошибок
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); err != nil {
		if oleErr, ok := err.(*ole.OleError); ok {
			logMessage("COM ошибка: код=%v, сообщение=%v", oleErr.Code(), oleErr.Error())
		} else {
			logMessage("Неизвестная ошибка: %v", err)
		}
		os.Exit(1)
	}

	defer ole.CoUninitialize()

	for {
		semaphore <- struct{}{} // Захватываем слот семафора

		// Запускаем через safeGoNoLog, т.к. го рутина переодически падает, не понятно из за API Windows или из за пакета go-ole
		safeGoNoLog(func() {
			defer func() { <-semaphore }() // Освобождаем слот после завершения

			// Проверка контекста перед выполнением основной логики
			if ctx.Err() != nil {
				logMessage("Получен сигнал о завершении работы приложения")
				return
			}

			time.Sleep(time.Second) // Задержка перед созданием COM объектов

			// Гибкое управление COM потоками || Локальная инициализация COM для каждой горутины
			comshim.Add(1)
			defer comshim.Done()

			// В основном цикле
			if !isOutlookRunning() {
				logMessage("Outlook не запущен. Попытка запуска...")
				if err := startOutlook(); err != nil {
					logMessage("Ошибка запуска Outlook: %v", err)
					return
				}

				if ctx.Err() != nil {
					logMessage("Получен сигнал о завершении работы приложения")
					return
				}

				time.Sleep(45 * time.Second) // Увеличенное время для инициализации
			}

			// Пытаемся инициализировать Outlook
			outlook, ns, err := initializeOutlook()
			if err != nil {
				logMessage("Ошибка инициализации Outlook: %v", err)

				// Завершаем процесс OUTLOOK.EXE
				err = killOutlookProcess()
				if err != nil {
					logMessage("Ошибка при завершении процесса OUTLOOK.EXE: %v", err)
				} else {
					logMessage("Попытка повторной инициализации Outlook после завершения процесса...")
				}
				return
			}

			folders := getTargetFolders(ns)
			if len(folders) == 0 {
				logMessage("Не найдено ни одной целевой папки")
				releaseObjects(outlook, ns)
				return
			}

			processFolders(folders)

			releaseObjects(outlook, ns)

		})

		// Ждем перед следующей попыткой
		waitNextCheck()
	}

}

func releaseObjects(objs ...*ole.IDispatch) {
	for _, obj := range objs {
		if obj != nil {
			obj.Release()
		}
	}
}

func isOutlookRunning() bool {
	processes, err := process.Processes()
	if err != nil {
		log.Printf("Ошибка получения списка процессов: %v", err)
		return false
	}

	for _, p := range processes {
		name, err := p.Name()
		if err == nil && strings.EqualFold(name, "OUTLOOK.EXE") {
			return true
		}
	}
	return false
}

func startOutlook() error {
	paths := []string{
		`C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE`,
		`C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE`,
		`C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE`,
		`C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE`,
	}

	for _, path := range paths {
		if _, err := os.Stat(path); err == nil {
			return exec.Command(path).Start()
		}
	}
	return exec.Command("outlook.exe").Start()
}

func initializeOutlook() (*ole.IDispatch, *ole.IDispatch, error) {
	// Попытка получить существующий экземпляр Outlook
	unknown, err := ole.GetActiveObject(CLSID_OutlookApp, IID_IDispatch)
	if err != nil {
		logMessage("Не удалось получить активный объект Outlook. Попытка создания нового...")
		unknown, err = oleutil.CreateObject("Outlook.Application")
		if err != nil {
			return nil, nil, fmt.Errorf("Ошибка создания объекта Outlook: %v", err)
		}
	}

	// Получение интерфейса IDispatch
	outlook, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, nil, fmt.Errorf("Ошибка получения интерфейса: %v", err)
	}

	// Получение пространства имен MAPI
	ns := oleutil.MustCallMethod(outlook, "GetNamespace", "MAPI").ToIDispatch()
	return outlook, ns, nil
}

func killOutlookProcess() error {
	// Получаем список всех запущенных процессов
	processes, err := process.Processes()
	if err != nil {
		return fmt.Errorf("Не удалось получить список процессов: %v", err)
	}

	for _, p := range processes {
		name, err := p.Name()
		if err != nil {
			continue
		}

		if name == "OUTLOOK.EXE" {
			logMessage("Завершение процесса OUTLOOK.EXE (PID: %d)...", p.Pid)
			if err := p.Kill(); err != nil {
				logMessage("Ошибка завершения процесса OUTLOOK.EXE (PID: %d): %v", p.Pid, err)
			} else {
				logMessage("Процесс OUTLOOK.EXE (PID: %d) успешно завершен.", p.Pid)
			}
		}
	}

	return nil
}

func getTargetFolders(ns *ole.IDispatch) map[string]*ole.IDispatch {
	folders := make(map[string]*ole.IDispatch)

	for _, folderCfg := range config.Folders {
		folder, err := getFolder(ns, folderCfg.Name)
		if err != nil {
			logMessage("Ошибка поиска папки %s: %v", folderCfg.Name, err)
			continue
		}
		folders[folderCfg.Name] = folder
	}

	return folders
}

func getFolder(ns *ole.IDispatch, name string) (*ole.IDispatch, error) {
	if name == "Входящие" {
		return oleutil.MustCallMethod(ns, "GetDefaultFolder", 6).ToIDispatch(), nil
	}
	return findFolderRecursive(ns, name)
}

func findFolderRecursive(parent *ole.IDispatch, target string) (*ole.IDispatch, error) {
	folders := oleutil.MustGetProperty(parent, "Folders").ToIDispatch()
	defer folders.Release()

	count := int(oleutil.MustGetProperty(folders, "Count").Val)
	for i := 1; i <= count; i++ {
		folder := oleutil.MustCallMethod(folders, "Item", i).ToIDispatch()
		currentName := oleutil.MustGetProperty(folder, "Name").ToString()

		if currentName == target {
			return folder, nil
		}

		subFolder, err := findFolderRecursive(folder, target)
		folder.Release()
		if err == nil {
			return subFolder, nil
		}
	}

	return nil, fmt.Errorf("папка '%s' не найдена", target)
}

func processFolders(folders map[string]*ole.IDispatch) {
	for folderName, folder := range folders {
		items := oleutil.MustCallMethod(folder, "Items").ToIDispatch()
		defer items.Release()

		filtered := oleutil.MustCallMethod(items, "Restrict", "[UnRead] = true").ToIDispatch()
		defer filtered.Release()

		count := int(oleutil.MustGetProperty(filtered, "Count").Val)
		if count == 0 {
			continue
		}

		// logMessage("Найдено %d новых сообщений в папке '%s'", count, folderName)

		for i := 1; i <= count; i++ {
			item := oleutil.MustCallMethod(filtered, "Item", i).ToIDispatch()
			processEmail(item, folderName)
			item.Release()
		}
	}
}

func processEmail(item *ole.IDispatch, folderName string) {
	// subjectID := oleutil.MustGetProperty(item, "Subject").ToString()
	entryIDVar, err := oleutil.GetProperty(item, "EntryID")
	if err != nil {
		logMessage("Ошибка получения EntryID: %v", err)
		return
	}
	entryID := entryIDVar.ToString()

	mutexMsg.Lock()
	defer mutexMsg.Unlock()
	if processedEmails[entryID] {
		// logMessage("Сообщение уже отправлено в Telegram: %s", subjectID)
		return
	}
	processedEmails[entryID] = true

	sender := ""
	senderName := oleutil.MustGetProperty(item, "SenderName").ToString()
	senderEmail := oleutil.MustGetProperty(item, "SenderEmailAddress").ToString()

	if senderName != "" && senderName != senderEmail {
		sender = (senderName + " <" + senderEmail + ">")
	} else {
		sender = (senderEmail)
	}

	subject := oleutil.MustGetProperty(item, "Subject").ToString()
	body := oleutil.MustGetProperty(item, "Body").ToString()

	var folderConfig Folder
	for _, f := range config.Folders {
		if f.Name == folderName {
			folderConfig = f
			break
		}
	}

	message := formatMessage(folderName, sender, subject, body, folderConfig.MessageLength)
	chatID := folderConfig.ChatID
	if chatID == "" {
		chatID = config.Telegram.DefaultChatID
	}

	if err := sendTelegramMessage(message, chatID); err != nil {
		processedEmails[entryID] = false
		logMessage("Ошибка отправки в Telegram: %v", err)
	} else {
		logMessage("Сообщение успешно отправлено в Telegram: %s", subject)
	}
}

func formatMessage(folder, sender, subject, body string, maxLength int) string {
	var msg strings.Builder

	// Экранируем специальные символы для HTML
	folder = html.EscapeString(folder)
	sender = html.EscapeString(sender)
	subject = html.EscapeString(subject)

	// Добавляем заголовки
	if config.Telegram.UseEmojis {
		msg.WriteString("📥 <b>Папка:</b> " + folder + "\n")
		msg.WriteString("👤 <b>Отправитель:</b> " + sender + "\n")
		msg.WriteString("📧 <b>Тема:</b> " + subject + "\n")
	} else {
		msg.WriteString("<b>Папка:</b> " + folder + "\n")
		msg.WriteString("<b>Отправитель:</b> " + sender + "\n")
		msg.WriteString("<b>Тема:</b> " + subject + "\n")
	}

	// Добавляем тело сообщения только если maxLength ≠ 0
	if maxLength != 0 {
		body = html.EscapeString(body)
		// Обрезаем тело если указана максимальная длина
		if maxLength > 0 {
			body = truncateByRunes(body, maxLength)
		}
		msg.WriteString("<i>Сообщение:</i>\n" + body)
	}

	// Проверяем, нужно ли обрезать текст до строки СutText
	cutString := config.CutText
	if len(cutString) != 0 {
		re := regexp.MustCompile("(?i)" + regexp.QuoteMeta(cutString))
		fullMessage := msg.String()
		index := re.FindStringIndex(fullMessage)
		if index != nil {
			// Обрезаем текст до найденного индекса
			fullMessage = fullMessage[:index[0]]
			msg.Reset()                  // Очищаем текущий Builder
			msg.WriteString(fullMessage) // Записываем обрезанный текст
		}
	}

	// Проверяем общую длину сообщения
	finalMessage := msg.String()
	return truncateByRunes(finalMessage, 4000)
}

func truncateByRunes(text string, maxRunes int) string {
	runes := []rune(text)
	if len(runes) <= maxRunes {
		return text
	}
	return string(runes[:maxRunes]) + "..."
}

func sendTelegramMessage(text, chatID string) error {
	url := fmt.Sprintf("https://api.telegram.org/bot%s/sendMessage", config.Telegram.BotToken)
	_, err := postJSON(url, map[string]interface{}{
		"chat_id":    chatID,
		"text":       text,
		"parse_mode": "HTML",
	})

	if err == nil {
		logMessage("Уведомление отправлено в чат %s", chatID)
	}
	return err
}

func postJSON(url string, data interface{}) ([]byte, error) {
	jsonData, err := json.Marshal(data)
	if err != nil {
		return nil, fmt.Errorf("ошибка маршалинга JSON: %v", err)
	}

	req, err := http.NewRequest("POST", url, bytes.NewBuffer(jsonData))
	if err != nil {
		return nil, fmt.Errorf("ошибка создания запроса: %v", err)
	}
	req.Header.Set("Content-Type", "application/json")

	resp, err := httpClient.Do(req) // ← Используем клиент с прокси
	if err != nil {
		return nil, fmt.Errorf("ошибка выполнения запроса: %v", err)
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("ошибка чтения ответа: %v", err)
	}

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("неверный статус код: %d, тело ответа: %s", resp.StatusCode, string(body))
	}

	return body, nil
}

func waitNextCheck() {
	interval := config.CheckIntervalSeconds
	if interval <= 0 {
		interval = 10
	}
	time.Sleep(time.Duration(interval) * time.Second)
}

// Создаёт http.Client с поддержкой HTTP/HTTPS/SOCKS5 прокси
func NewHTTPClientWithProxy(cfg ProxyConfig) (*http.Client, error) {

	// Логирование статуса прокси
	if cfg.Enabled {
		authInfo := ""
		if cfg.Username != "" {
			authInfo = " (с авторизацией)"
		}
		logMessage("Прокси включен: %s://%s:%d%s",
			cfg.Type, cfg.Host, cfg.Port, authInfo)
	} else {
		logMessage("Прокси выключен: используется прямое подключение")
	}

	// Если прокси выключен — возвращаем обычный клиент
	if !cfg.Enabled {
		return &http.Client{Timeout: 30 * time.Second}, nil
	}

	// Создаём базовый транспорт
	transport := &http.Transport{
		TLSHandshakeTimeout:   10 * time.Second,
		ResponseHeaderTimeout: 10 * time.Second,
		ExpectContinueTimeout: 1 * time.Second,
	}

	// Настройка прокси в зависимости от типа
	switch cfg.Type {
	case "http", "https":
		// === HTTP/HTTPS прокси ===
		// Используем стандартный механизм http.Transport.Proxy
		proxyURL := &url.URL{
			Scheme: "http", // Всегда "http" для HTTP-прокси, даже если целевой сайт https
			Host:   fmt.Sprintf("%s:%d", cfg.Host, cfg.Port),
		}

		// Добавляем авторизацию, если есть
		if cfg.Username != "" {
			proxyURL.User = url.UserPassword(cfg.Username, cfg.Password)
		}

		// Устанавливаем прокси в транспорт
		transport.Proxy = http.ProxyURL(proxyURL)

		// Настройка обхода прокси (no_proxy)
		if len(cfg.NoProxy) > 0 {
			transport.Proxy = func(req *http.Request) (*url.URL, error) {
				// Проверяем список no_proxy
				for _, host := range cfg.NoProxy {
					if req.URL.Hostname() == host || net.ParseIP(req.URL.Hostname()).IsLoopback() {
						return nil, nil // Не использовать прокси
					}
				}
				// Иначе — использовать прокси
				return proxyURL, nil
			}
		}

		logMessage("HTTP-прокси настроен: %s:%d", cfg.Host, cfg.Port)

	case "socks5":
		// === SOCKS5 прокси ===
		auth := (*proxy.Auth)(nil)
		if cfg.Username != "" {
			auth = &proxy.Auth{
				User:     cfg.Username,
				Password: cfg.Password,
			}
		}

		dialer, err := proxy.SOCKS5("tcp",
			fmt.Sprintf("%s:%d", cfg.Host, cfg.Port),
			auth,
			proxy.Direct)
		if err != nil {
			logMessage("Ошибка настройки SOCKS5-прокси: %v", err)
			return nil, fmt.Errorf("ошибка настройки SOCKS5-прокси: %w", err)
		}

		// Устанавливаем кастомный dialer в транспорт
		transport.DialContext = func(ctx context.Context, network, addr string) (net.Conn, error) {
			// Обход прокси для no_proxy
			host, _, err := net.SplitHostPort(addr)
			if err == nil {
				for _, noProxy := range cfg.NoProxy {
					if host == noProxy || net.ParseIP(host).IsLoopback() {
						return (&net.Dialer{}).DialContext(ctx, network, addr)
					}
				}
			}
			return dialer.Dial(network, addr)
		}

		logMessage("SOCKS5-прокси настроен: %s:%d", cfg.Host, cfg.Port)

	default:
		err := fmt.Errorf("неподдерживаемый тип прокси: %s", cfg.Type)
		logMessage("%v", err)
		return nil, err
	}

	logMessage("HTTP-клиент инициализирован (таймаут: 30с)")

	return &http.Client{
		Transport: transport,
		Timeout:   30 * time.Second,
	}, nil
}
