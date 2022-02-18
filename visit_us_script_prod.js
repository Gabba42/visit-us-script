//current available days as of Jan2022: Sunday, Tuesday, Thursday
//nmo starts 30 mins after tours (nmo and other time have the same time) 

var form = FormApp.getActiveForm(); 

//get date/time selection sections from active form
var tourDateTime = form.getItemById(1782567039); 
var nmoDateTime = form.getItemById(158641723); 
var otherDateTime = form.getItemById(1019984296); 

//get date/time selections items as a list item so we can add choices 
var tourDateTimeAsListItem = tourDateTime.asListItem(); 
var nmoDateTimeAsListItem = nmoDateTime.asListItem(); 
var otherDateTimeAsListItem = otherDateTime.asListItem(); 

//returns true if date is today; false otherwise
function isToday(date) {
  var today = new Date(); 
  return date.getDate() === today.getDate() && 
         date.getMonth() === today.getMonth() && 
         date.getFullYear() === today.getFullYear(); 
}

//returns true if date is tomorrow; false otherwise
function isTomorrow(date) {
  var tomorrow = new Date(); 
  return date.getDate() === tomorrow.getDate()+1 && 
         date.getMonth() === tomorrow.getMonth() && 
         date.getFullYear() === tomorrow.getFullYear(); 
}

  //returns a result of true if the date to check is a holiday 
  function isHoliday(date) {

    let result = false; 

    const holidaysInYear = {
        christmasVaca1: new Date("December 30, 2021"),
        christmasVaca2: new Date("January 2, 2022"),
        christmasVaca3: new Date("January 4, 2022"),
        christmasVaca4: new Date("January 6, 2022"),
        christmasVaca5: new Date("January 9, 2022"),
        christmasVaca6: new Date("January 11, 2022"),
        mlkDay: new Date("January 17, 2022"),
        memorialDay: new Date("May 30, 2022"),
        independenceDay: new Date("July 4, 2022"),
        laborDay: new Date("September 4, 2022"), 
        veteransDay: new Date("November 11, 2022"), 
        thanksgivingEve: new Date("November 23, 2022"),
        thanksgivingDay: new Date("November 24, 2022"), 
        christmasEve: new Date("December 24, 2022"), 
        christmasDay: new Date("December 25, 2022") 
    };
    
    Object.values(holidaysInYear).every(holiday => {
    
        if (date.getDate() === holiday.getDate() 
            && date.getMonth() === holiday.getMonth () 
            && date.getFullYear() === holiday.getFullYear()) {
            
            //if true, return false so it can "break" out of the loop - needed for every() function
            result = true; 
            return false; 
            
        } 
        else {
            //return true so that it can "continue"
            return true; 
        }
    }); 

    return result; 

  }

//get all available days
function getAvailableDays() {
  var date = new Date(); 
  var days = []; 
  for(let i = 0; i < 6;  i++) { //because we only want two weeks worth of options 
    if(date.getDay() === 0 || date.getDay() === 2 || date.getDay() === 4) { //CHANGE DATES HERE!!
      if(!isToday(date) && !isTomorrow(date) && !isHoliday(date)){ //ADD DATES TO EXCLUDE!! 
          days.push(new Date(date));
      }
      else{
        i-=1;
      }
    }
    else {
      i-=1;
    }
    date.setDate(date.getDate() +1); 
  }
  return days; 
}

//format dates and add times for tours
function formatTourDateAndTime(dateArr) {
  var formattedDateArr = []; 
  for(let i = 0; i < dateArr.length; i++) {
    if(dateArr[i].getDay() === 0) { 
      formattedDateArr.push(dateArr[i].toDateString() + ' @ 2:00PM (MST)');
    }
    else if(dateArr[i].getDay() === 2) {
      formattedDateArr.push(dateArr[i].toDateString() + ' @ 7:00AM (MST)');
    }
    else {
      formattedDateArr.push(dateArr[i].toDateString() + ' @ 5:00PM (MST)');
    }
  }
  return formattedDateArr; 
}

//format dates and add times for new member orientation
function formatNmoOtherDateAndTime(dateArr) {
  var formattedDateArr = []; 
  for(let i = 0; i < dateArr.length; i++) {
    if(dateArr[i].getDay() === 0) { 
      formattedDateArr.push(dateArr[i].toDateString() + ' @ 2:30PM (MST)');
    }
    else if(dateArr[i].getDay() === 2) {
      formattedDateArr.push(dateArr[i].toDateString() + ' @ 7:00AM (MST)');
    }
    else {
      formattedDateArr.push(dateArr[i].toDateString() + ' @ 5:30PM (MST)');
    }
  }
  return formattedDateArr; 
}

//creates date/time choices in form for tours 
function setDateTimeChoicesForTours() {
  var dateArray = getAvailableDays();
  var formattedDateAndTimeArr = formatTourDateAndTime(dateArray);
  tourDateTimeAsListItem.setChoiceValues(formattedDateAndTimeArr); 
}

//creates date/time choices in form for nmo & other 
function setDateTimeChoicesForNmoOther() {
  var dateArray = getAvailableDays();
  var formattedDateAndTimeArr = formatNmoOtherDateAndTime(dateArray);
  nmoDateTimeAsListItem.setChoiceValues(formattedDateAndTimeArr);
  otherDateTimeAsListItem.setChoiceValues(formattedDateAndTimeArr); 
}

//main function to call 
function setDateTimeChoices() {
  setDateTimeChoicesForTours(); 
  setDateTimeChoicesForNmoOther();
}

Logger.log(getAvailableDays()); //making sure we are getting the available days 
