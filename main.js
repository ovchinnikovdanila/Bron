function ExpInCal() {
  
  var table = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dl = table.getLastRow();
  var data = table.getRange("A2:H"+dl).getValues();
  var k = data.length - 1;

  if (!!data[k][3]) {
    
    var triggerForEv = 1;

    switch(data[k][2]) {
      case "Конференц-зал":
        var cal = CalendarApp.getOwnedCalendarById('799lageq18e131aelcqct9us2c@group.calendar.google.com');
        var events = cal.getEvents(new Date(2021,01,01),new Date(2030,01,01));
        break;
      case "Аудитория 206":
        var cal = CalendarApp.getOwnedCalendarById('t14er1ojohcgb4dimffg15i8hc@group.calendar.google.com');
        var events = cal.getEvents(new Date(2021,01,01),new Date(2030,01,01));
        break;
      case "Аудитория 312":
        var cal = CalendarApp.getOwnedCalendarById('2gst2kl96dcqe6iedg37h490kk@group.calendar.google.com');
        var events = cal.getEvents(new Date(2021,01,01),new Date(2030,01,01));
        break;
      case "Аудитория 410":
        var cal = CalendarApp.getOwnedCalendarById('kls9qt5j4s78aucc9o8dfp3tnk@group.calendar.google.com');
        var events = cal.getEvents(new Date(2021,01,01),new Date(2030,01,01));
        break;
      case "Читальный зал №1":
        var cal = CalendarApp.getOwnedCalendarById('4m7nh9d57tb0stts1u653u7m3k@group.calendar.google.com');
        var events = cal.getEvents(new Date(2021,01,01),new Date(2030,01,01));
        break;
      case "Читальный зал №2":
        var cal = CalendarApp.getOwnedCalendarById('spckas0i7filpro02rf54rnj40@group.calendar.google.com');
        var events = cal.getEvents(new Date(2021,01,01),new Date(2030,01,01));
        break;
      default:
        triggerForEv = 0;
    }

    if (triggerForEv == 1) {

      var nach = String(data[k][0]);
      var kon = String(data[k][1]);
      var trigger = 0;
      
      if ((nach.getFullYear != kon.getFullYear) || (nach.getMonth != kon.getMonth) || (nach.getDay != kon.getDay) || (data[k][1] < data[k][0])) {
        console.log("dataerr")
        MailApp.sendEmail(data[k][5], "Бронирование "+ data[k][2] + " " + data[k][3], "Уважаемый/ая " + data[k][4] +"! Время не было забронировано, неверно введены даты начала и окончания!");
      } else {

        for (var j = 0 ; j < events.length; j++) {

          var nachEv = String(events[j].getStartTime());
          var konEv = String(events[j].getEndTime());
            
          if (((data[k][0] > events[j].getStartTime()) && (data[k][0] < events[j].getEndTime())) || ((data[k][1] > events[j].getStartTime()) && (data[k][1] < events[j].getEndTime())) || (nach == nachEv) || (kon == konEv) || ((data[k][0] < events[j].getStartTime()) && (data[k][1] > events[j].getEndTime())))
            trigger = 1;
            
        }

        if (trigger == 1) {
          console.log("dubl"); 
          MailApp.sendEmail(data[k][5], "Бронирование "+ data[k][2] + " " + data[k][3],"Уважаемый/ая " + data[k][4] + "! Мероприятие не было забронированно. Время уже занято.");
          }

        if (trigger == 0) {

          var calEvents = CalendarApp.getOwnedCalendarById('p6i7g8casck7g8iv0lch1eu9f8@group.calendar.google.com');
          calEvents.createEvent(data[k][2] + " " + data[k][3], data[k][0], data[k][1], {description: "Мероприятие забранировал(а):" + data[k][4] + ". e-mail: " + data[k][5]});
          cal.createEvent(data[k][3], data[k][0], data[k][1], {description: "Мероприятие забранировал(а):" + data[k][4] + ". e-mail: " + data[k][5]});

          console.log("Создано");
          MailApp.sendEmail(data[k][5], "Бронирование " +  data[k][2] + " " + data[k][3], "Уважаемый/ая " + data[k][4] + ". Ваше мероприятие успешно забронировано!");

        }
      
      }
    
    }
  
  }

}
