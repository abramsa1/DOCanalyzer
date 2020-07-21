// этот js файл содержит настройки жизненного цикла приложения
// деструктурируем необходимые компоненты из пакета electron
const { app, BrowserWindow, ipcMain, dialog } = require('electron');
// получаем методы для работы с директориями
const path = require('path');

// отлавливаем метку open-file-dialog из render.js для выбора директории и вызываем анонимную функцию
ipcMain.on('open-file-dialog', async (event) => {
  // открываем окно выбора папки
  const selectedPath = await dialog.showOpenDialog({
    // устанавливаем свойство на выбор только папок, а не файлов
    properties: ['openDirectory'],
  });
  // отправляем сообщение в render скрипт, с меткой выбора папки и данными - адресом выбранной директории
  event.sender.send('selected-directory', selectedPath.filePaths[0]);
});

// функция создания окна приложения
function createWindow() {
  // создаем окно браузера
  let win = new BrowserWindow({
    // устанавливаем ширину и высоту окна приложения в пикселях
    width: 530,
    height: 730,
    // устанавливаем путь до иконки приложения
    icon: path.join(__dirname, '/img/icon.ico'),
    // скрываем классическое меню
    autoHideMenuBar: true,
    // отключаем возможность изменять размер окна приложения
    resizable: false,
    // свойство - объект с настройками приложения
    webPreferences: {
      // включаем поддержку скриптов node.js
      nodeIntegration: true,
    },
  });
  // загружаем файл html для отображения интерфейса приложения
  win.loadFile('index.html');
}

// вызываем функцию создания окна приложения, когда electron закончил инициализацию
app.whenReady().then(createWindow);
