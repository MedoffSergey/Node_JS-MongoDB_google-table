//Подключение модулей
const express = require('express');         //  Подключаем Express
const app = express();                      //  Инциализируем Express в app
const path = require('path');               //  Модуль path позволяет указывать пути к дирректориям
const pug = require('pug');                 //  Шабланизатор Pug
const moment = require('moment');           //  Moment js
require('moment-weekday-calc');             //  дополнение (для подсчета рабочих дней)

const GoogleSpreadsheet = require('google-spreadsheet');    //  Модуль для получения данных с Гугловой таблицы
const { promisify } = require('util')                       //  Модуль для получения данных с Гугловой таблицы
const creds = require('./client_secret.json');              //  Серетный ключ клиента гугловой таблицы

const MongoClient = require('mongodb').MongoClient;         // Модуль MongoClient подключимся как клиент
const assert = require('assert');
const url = 'mongodb://localhost:27017';                    // Соединение URL
const mongoClient = new MongoClient(url, {
  useNewUrlParser: true
});
const dbName = 'Foremans'; // Название базы данных

app.set('views', path.join(__dirname, 'views'));  //указываем путь к pug файлам
app.set('view engine', 'pug');                    // указываем используемый шаблонизатор HTML кода

app.use(express.static(path.join(__dirname, 'public')));    //добовляет файлы которые на компьютере для загрузки если они имеются

const correctColumnNames = [        //  Массив с правильными маршрутами
  'surname',                        //  Фамилия
  'name',                           //  Имя
  'patronymic',                     //  Отчество
  'employment_date',                //  Дата приема на работу
  'date_of_dismissal',              //  Дата увольнения
  'number_of_working_days',         //количество рабочих дней
  'number_of_meetings',             //  Количество слетов
  'the_number_of_successful_objects',//  Количество успешных обьектов
  'total_objects',                  //  Всего объектов
  'rates_percent',                  //  Процент Слетов
  'information_upload_date',        //  Дата загрузки информации
]


function showForemanTables(foremans) { // получаем массив с Соответствовающими именами
  const user = {};

  for (let key in foremans) {
    if (key.match(/[a-zA-Z]/)) {      //  Используем регулярное выражение для поиска лишних данных
      delete foremans[key]            //  Убираем лишние данные на Латынице
    }
  }

  function number_of_working(employ,dismissal){                         //функция подсчета рабочих дней (без учета выходных)
    const	firstDate = moment(user[correctColumnNames[3]], 'DD-MM-YYYY')
    const lastDate = moment(user[correctColumnNames[4]], 'DD-MM-YYYY')
    const diff = (moment().isoWeekdayCalc(firstDate,lastDate,[1,2,3,4,5]));
    if(diff){
      user[correctColumnNames[5]] = diff;
    }
  }

  Object.keys(foremans).forEach((item, index) => {
    let fullNameArr = [];

    if(item=='прораб'){
        let fullNameArr = foremans[item].split(' ')
        user[correctColumnNames[0]] = fullNameArr[0]
        user[correctColumnNames[1]] = fullNameArr[1]
        user[correctColumnNames[2]] = fullNameArr[2]
    }

    if(item=='датаприеманаработу'){
      user[correctColumnNames[3]] = foremans['датаприеманаработу']
    }
    if(item=='датаувольнения'){
      user[correctColumnNames[4]] = foremans['датаувольнения']
    }
    if(item=='количествослетов'){
      user[correctColumnNames[6]] = foremans['количествослетов']
    }
    if(item=='количествоуспешныхобъектов'){
      user[correctColumnNames[7]] = foremans['количествоуспешныхобъектов']
    }
    number_of_working(user[correctColumnNames[3]],user[correctColumnNames[4]])  //вызов функции подсчета рабочих дней
    user[correctColumnNames[8]] = Number(user[correctColumnNames[7]]) + Number(user[correctColumnNames[6]])                     // Всего объектов
    user[correctColumnNames[9]] = (Number(user[correctColumnNames[7]]) / Number(user[correctColumnNames[6]])).toFixed(1)+'%';   // % Слетов
    const today = new Date();
    user[correctColumnNames[10]] = moment().format('YYYY-MM-DD HH:mm:ss');                // Дата загрузки информации
  })
  return user
}

async function accessSpreadsheet() {                                                 // Функция стягивает данные с гугловой таблицы
  const doc = new GoogleSpreadsheet('1F_tXXdFUO60xfF4kX5qmSWA-gxq_BM39D6VEMmoStKA'); // Подключаем гугл таблицу
  await promisify(doc.useServiceAccountAuth)(creds);                                 // Вводим  секретный json с гугл таблицы
  const info = await promisify(doc.getInfo)();
  const sheet = info.worksheets[0];
  const rows = await promisify(sheet.getRows)({
    offset: 1
  })

  var newArr = rows.map(function(elem) {  //переберем rows
    let tryTables = showForemanTables(elem)
    return tryTables
  });
  return newArr
}


//Главная страница
app.get('/', function(req, res) {
  mongoClient.connect(function(err, client){
      let mongodbList
      const db = client.db(dbName);
      const collection = db.collection("Foremans");

      if(err) return console.log(err);

      collection.find().toArray(function(err, results){
          mongodbList = results;

          res.render('index', {   //рендерим файл index.pug
            title: 'Show Tables', //титл страницы
            list: mongodbList     //сюда засовываем список как дополнительное свойство
          });
      });
  });
});

app.get('/showForeman', function(req, res) {
  accessSpreadsheet().then(resultArr => {
    mongoClient.connect(function(err, client) { //Используем метод подключения для подключения к серверу
      const db = client.db(dbName);
        db.collection("Foremans").drop()  // очистим бд

      const collection = db.collection("Foremans");
      collection.insertMany(resultArr, function(err, result) {
        if (err) {
          return console.log(err);
        }
        console.log("данные успешно заполнены")
          res.json({success:1}) // отправим ответ
      });
    });
  });
});

app.get('/search', function(req, res) { // обработка поиска и отображение полученных результатов
  let searchResult = req.query.search; //присваиваем переменной результат запроса клиента

  mongoClient.connect(function(err, client){
      let mongodbList
      const db = client.db(dbName);
      const collection = db.collection("Foremans");

      if(err) return console.log(err);

      collection.find({$or: [{"name": searchResult}, {"surname": searchResult}, {"patronymic": searchResult}]}).toArray(function(err, results){
          mongodbList = results;

          console.log(mongodbList)
          res.render('index', {   //рендерим файл index.pug
            title: 'Search', //заголовок
            list: mongodbList     //сюда пихаем список как дополнительное свойство
          });
      });
  });
})

//запускаем сервер
app.listen(3000, function() {
  console.log('Отслеживаем порт: 3000!');
});
