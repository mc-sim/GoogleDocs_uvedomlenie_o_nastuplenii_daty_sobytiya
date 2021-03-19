function emailsend () //Отправляет e-mail при выполнении условия описанного в if
{
  var sovpadenieDat = in_array(reminder (), perebor_dat());
  if(sovpadenieDat > -1) 
  {
    MailApp.sendEmail(
                      "your_mail_here@mail.ru",
                       perebor_dat()[in_array(reminder (), perebor_dat())] + " закончится срок действия " + vozvrat_temy()[in_array(reminder (), perebor_dat())],                             
                      'Привет! Заканчивается срок действия ' + vozvrat_temy()[in_array(reminder (), perebor_dat())]
                     )
  }
} 

function reminder ()  //Прибавляет к текущей дате 14 дней. Возвращается дата в формате DD.MM.YYYY для сравнения с содержимым массива perebor_dat() 
{
eval(UrlFetchApp.fetch('https://momentjs.com/downloads/moment.js').getContentText()); //подключаем библиотеку Moment для правильного форматирования даты
var result = new Date();
var result = result.setDate(result.getDate() + 14);
var result = moment(result).format('DD.MM.YYYY');
return result; 
}

function perebor_dat() //Возвращает массив отформатированных значений содержащихся на текущем листе в диапазоне B2:B20 (закомментированная строка - A2:Z2. Проблема с возвратом значений из ячеек AA2, AB2 так как цикл проходит только по алфавиту)
{
  var kolonkaplus = [];
  for (var i=3; i<=21; i++)
  {
//    kolonkaplus.push(Utilities.formatDate(new Date(SpreadsheetApp.getActiveSheet().getRange(String.fromCharCode(i)+'2').getValues()), "GMT+3", "d.M.yyyy"));
kolonkaplus.push(Utilities.formatDate(new Date(SpreadsheetApp.getActiveSheet().getRange(String.fromCharCode(66)+(i)).getValues()), "GMT+3", "dd.MM.yyyy"));
  }
  return kolonkaplus;
}

function vozvrat_temy() //Возвращает массив отформатированных значений содержащихся на текущем листе в диапазоне A2:A20 (закомментированная строка -  A1:Z1. Проблема с возвратом значений из ячеек AA1, AB1 так как цикл проходит только по алфавиту
{
  var kolonkaplus = [];
  for (var i=2; i<=19; i++)
  {
    //kolonkaplus.push(SpreadsheetApp.getActiveSheet().getRange(String.fromCharCode(i)+'1').getValues());
    kolonkaplus.push(SpreadsheetApp.getActiveSheet().getRange(String.fromCharCode(65)+(i)).getValues());
  }
  return kolonkaplus;
}

function in_array(value, array) //Проверяет массив на наличие в нем значения
{
  for(var i = 0; i < array.length; i++) 
  {
    if(array[i] == value) return i;
  }
  return -1;
}
