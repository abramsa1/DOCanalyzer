// деструктурируем необходимые компоненты из пакета electron
const { shell, ipcRenderer } = require('electron');
// получаем компонент textract из загружаемого пакета 
const textract = require('textract');
// получаем компонент WordExtractor из загружаемого пакета 
const WordExtractor = require('word-extractor');
// инициализируем новый класс
const extractor = new WordExtractor();
// получаем компоненты для работы с директориями
const dirPath = require('path');
// получаем компоненты для работы с файловой системой
let fs = require('fs');

// Получаем DOM элементы
const selectDirBtn = document.getElementById('select-directory'); // кнопка выбора папки
const out = document.getElementById('output'); // поле вывода данных
const key = document.getElementById('key'); // поле ввода ключевой фразы
const docCount = document.getElementById('documentCount'); // поля вывода количества документов в папке

const processBar = document.querySelector('.progress__bar'); // прогресс бар
const wrapperBar = document.querySelector('.wrapper-bar'); // обертка прогресс бара
const textBar = document.querySelector('.textBar'); // подпись к прогресс бару

let stepCount = 0; // шаг прогресс бара
let procentCount = 0; // текущее значение прогресс бара
let userPath = ''; // текущая выбранная директория

// с помощью делегирования - прослушиваем событие - клик левой кнопкой мыши
// по клику открываем директорию нужного документа
out.addEventListener('click', (event) => {
  // проверяем, кликнули ли по нужному элементу
  let elem = event.target.closest('div.document');
  // если да, то открываем проводник
  if (elem) shell.showItemInFolder(`${userPath}\\${elem.dataset.name}`);
});

// по клику на кнопку отправляем сообщение в файл жизненного цикла для выбора директории
selectDirBtn.addEventListener('click', () => {
  ipcRenderer.send('open-file-dialog');
});

// ожидаем метку 'selected-directory' из файла жизненного цикла
// если метка с данными (адрес директории) пришла - вызываем функцию 
ipcRenderer.on('selected-directory', (event, path) => {
  // запоминаем путь в глобальную переменную
  userPath = path;
  // в качестве текста на кнопке устанавливаем адрес директории, если адрес не выбран - значение по умолчанию 'Выберите папку'
  selectDirBtn.textContent = userPath || 'Выберите папку';
  // вызываем основную функцию
  main(path);
});

// объект данных уведомления
const notification = {
  // заголовок уведомления
  title: '.DOC Analyzer',
  // текст уведомления
  body: 'Анализ файлов завершен',
  // иконка
  icon: dirPath.join(__dirname, '/img/icon.png'),
};

// функция вызывает всплывающее уведомление windows 
function showNotification() {
  // создаем уведомление через методы electron
  const myNotification = new window.Notification(notification.title, notification);
}

// основная функция
// в аргументе - выбранный пользователем путь (адрес директории)
async function main(path) {
  // очищаем поле для вывода данных
  out.textContent = '';
  // получаем значение из поля ввода ключевой фразы, если фразу нет, выбираем строку по умолчанию - 'Введите фразу';
  // переводим ключевую фразу в нижний регистр
  let keyPhrase = (key.value || 'Введите фразу').toLowerCase();
  // получаем массив имен .doc элементов в директории 
  let doc = getFilesFromPath(path, 'doc');
  // получаем массив имен .docx элементов в директории 
  let docx = getFilesFromPath(path, 'docx');
  // если оба массива пустые - выводим сообщение об отсутствии файлов и выходим из функции
  if (!(doc.length || docx.length)) {
    renderZeroMessage();
    return;
  }
  // инициализируем прогресс бар, отправляем ему общее количество найденных документов
  progress(doc.length + docx.length);

  // с помощью оператора spread (...) создаем результирующий массив объектов
  let arrDocs = [
    ...(await getFilesArr(doc, path, keyPhrase, 'doc')),
    ...(await getFilesArr(docx, path, keyPhrase, 'docx')),
  ];
  // сортируем массив по убыванию количества повторений ключевой фразы
  arrDocs.sort((prev, next) => next.count - prev.count);
  // анализ завершен - скрываем прогресс бар
  wrapperBar.classList.add('hide');
  // вызываем функцию для отображения результата анализа 
  renderFileNames(arrDocs);
  // вызываем функцию показа уведомления 
  showNotification();
}

// функция возвращает массив имен элементов из указанной директории
// в аргументе - путь (директория) и формат файла
function getFilesFromPath(path, format) {
  // методом node.js получаем коллекцию элементов из директории 
  let files = fs.readdirSync(path);
  // инициализируем регулярное выражение для фильтрации файлов согласно параметру format
  let reg = new RegExp(`^([^~$]).*(\\.${format})$`, 'i');
  // возвращаем массив элементов, которые соответствуют регулярному выражению
  return files.filter((elem) => elem.match(reg)); //docx
}

// функция формирует массив объектов которые содержат данные об имени файла и количестве повторений ключевой фразы
// в аргументе - массив документов из указанной директории, директория, ключевая фраза и формат файлов
async function getFilesArr(arr, path, keyPhrase, format) {
  let countValue = 0;
  let newArr = [];
  // прогоняем через цикл все элементы массива
  for (let elem of arr) {
    // вызываем функцию для отображения прогресс бара
    progressStep(elem);
    // вызываем функцию-счетчик для документов doc
    if (format === 'doc') countValue = await docAnalys(`${path}\\${elem}`, keyPhrase);
    // вызываем функцию-счетчик для документов docx
    else countValue = await docxAnalys(`${path}\\${elem}`, keyPhrase);
    // если результат счетчика с ошибкой, пропускаем этот документ
    if (countValue === undefined) continue;
    // добавляем в массив объект с именем файла и колличестве повторений ключевой фразы
    newArr.push({ filename: elem, count: countValue });
  }
  // возвращаем массив с данными
  return newArr;
}

// функция для анализа docx файлов
// в аргументе - полный путь до документа и ключевая фраза 
function docxAnalys(name, key) {
  // в этой функции воспользуемся методом fromFileWithPath из загружаемого пакета textract
  // т.к. этот метод асинхронный, будем возвращать обещание (promise)
  return new Promise((resolve) => {
    textract.fromFileWithPath(name, (err, text) => {
      // если возникнет ошибка выйдем из функции, ничего не возвращая
      if (err) resolve();
      // иначе, возвращаем количество повторений ключевой фразы в документе
      resolve(getCount(text, key));
    });
  });
}

// функция для анализа doc файлов
// в аргументе - полный путь до документа и ключевая фраза 
function docAnalys(name, key) {
  // получаем результат метода extract, куда отправляем полный путь до файла doc 
  // этот метод из загружаемого пакета extractor, возвращает обещание (promise) 
  const extracted = extractor.extract(name);
  // возвращаем количество повторений ключевой фразы в документе
  return extracted
    // т.к. extractor возвращает обещание, воспользуемся методом then
    .then((doc) => {
      // получаем текст документа в переменную
      let text = doc.getBody();
      // возвращаем количество повторений ключевой фразы в документе
      return getCount(text, key); 
    })
    // catch будет вызван в случае ошибки в обещании в методе extrac
    .catch((err) => {
      // выводим ошибку в консоль разработчика
      console.error(err);
      // выходим из функции
      return;
    });
}

// функция подсчитывает количество повторений ключевой фразы в тексте
// в аргументе - текст, полученный из doc файла и ключевая фраза поиска - ключ
function getCount(text, key) {
  // с помощью регулярного выражения удаляем из текста все пустые строки и переводим символы в нижний регистр
  text = text.replace(/\n+/g, ' ').toLowerCase();
  // инициализируем переменную счетчик
  let counter = 0;
  // вычисляем начальную позицию где встречается ключ
  let pos = text.indexOf(key, 0);
  // если ключ присутствует - запускаем цикл в котором подсчитываем количество позиций
  while (pos !== -1) {
    pos = text.indexOf(key, pos + 1);
    counter++;
  }
  // возвращаем счетчик
  return counter;
}

// функция для отображения найденных документов
function renderFileNames(arr) {
  let items = '';
  // выводим количество найденных документов
  docCount.textContent = arr.length;
  // генерируем html код для отображения найденных документов
  for (let elem of arr) {
    items += `
    <div data-name="${elem.filename}" class="document">
    <img class="document-logo" src="./img/document.png" alt="document-logo">
    <div class="text-block">
      <p class="document-name">${elem.filename}</p>
      <p class="words-count">Повторений фразы: <span>${elem.count}</span></p>
    </div>
    </div>
    `;
  }
  // добавляем сообщение в html верстку
  out.insertAdjacentHTML('afterbegin', items);
}

// функция отображения сообщения об отсутствии документов в директории
function renderZeroMessage() {
  // обнуляем количество найденных документов
  docCount.textContent = 0;
  // генерируем сообщение об отсутствии сообщений
  let item =
    '<h3 class="welcome-header">В указанной директории файлы Microsoft Word отсутствуют.</h3>';
  // добавляем сообщение в html верстку
  out.insertAdjacentHTML('afterbegin', item);
}

// функция инициализации прогресс бара
// в качестве аргумента - количество анализируемых документов
function progress(elemCount) {
  // обнуляем текущее значение прогресс бара
  procentCount = 0;
  // вычисляем шаг прогресс бара
  stepCount = 100 / elemCount;
  // отображаем прогресс бар - удаляем класс hide 
  wrapperBar.classList.remove('hide');
}

// функция, устанавливающая процент анализа документов
// в качестве аргумента принимаем имя текущего документа который анализируем
function progressStep(elem) {
  // к текущему значению прогресс бара прибавляем шаг загрузки
  procentCount += stepCount;
  // если значение прогресс бара > 100 выходим из функции
  if (procentCount >= 100) return;
  // устанавливаем ширину элемента html который эмулирует прогресс бар на текущее значение прогресс бара
  processBar.style.width = `${procentCount}%`;
  // подписываем прогресс бар - устанавливаем имя документа который анализируем
  textBar.textContent = elem;
}
