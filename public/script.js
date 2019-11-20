'use strict'

function showForeman() {
  $.ajax({
    url: "/showForeman", //путь
    type: "GET",         //Метод отправки
    success: function(data) {
      if(data.success){
        window.location.reload() //если успешно, то перезапускаем страницу через аякс
      }
      else console.log("error script.js")
    }
  });
}
