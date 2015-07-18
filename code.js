function emailsend () //Отправляет e-mail при выполнении условия описанного в if
{
  var sovpadenieDat = in_array(reminder (), perebor_dat());
  if(sovpadenieDat > -1) 
  {
    MailApp.sendEmail(
                      "1777510@mail.ru",
                       perebor_dat()[in_array(reminder (), perebor_dat())] + " пройдет рассылка " + vozvrat_temy()[in_array(reminder (), perebor_dat())],                             
                      'Привет! Поставь пожалуйста в этот график: https://docs.google.com/spreadsheets/d/1TLY22BHVyzA4kLPyG3eQ4Phia2O675Vuywj9ZseVeOI/edit#gid=12 \r\nсвои акции релевантные теме рассылки.'
                     )
  }
} 
 
function reminder ()  //Прибавляет к текущей дате 3 дня. Возвращается не дата, а строка но ее сравнение с содержимым массива perebor_dat() прокатывает
{
  var dates = new Date()
  var reminderMonth = dates.getMonth()+1;
  var reminderDate = dates.getDate() +3 ;
  var reminderYear = '2015';
  return reminderDate + "." + reminderMonth 	+ "." + reminderYear;
}

function perebor_dat() //Возвращает массив отформатированных значений содержащихся на текущем листе в диапазоне A2:Z2. Проблема с возвратом значений из ячеек AA2, AB2 так как цикл проходит только по алфавиту
{
  var kolonkaplus = [];
  for (var i=65; i<=90; i++)
  {
    kolonkaplus.push(Utilities.formatDate(new Date(SpreadsheetApp.getActiveSheet().getRange(String.fromCharCode(i)+'2').getValues()), "GMT+3", "d.M.yyyy"));
  }
  return kolonkaplus;
}

function vozvrat_temy() //Возвращает массив отформатированных значений содержащихся на текущем листе в диапазоне A1:Z1. Проблема с возвратом значений из ячеек AA1, AB1 так как цикл проходит только по алфавиту
{
  var kolonkaplus = [];
  for (var i=65; i<=90; i++)
  {
    kolonkaplus.push(SpreadsheetApp.getActiveSheet().getRange(String.fromCharCode(i)+'1').getValues());
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