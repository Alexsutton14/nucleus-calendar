const personId = /* paste personID here */
const expires = "/* paste cookie expiry here */"
const cookie = "/* paste cookie here */"
const resourceId = /* paste resourceID here */

const options = {
  useDefaultCalendar: false,
  calendarName: "Nucleus"
}

//
// DON'T EDIT BELOW THIS
//

let userProps = PropertiesService.getUserProperties();

let session;
let requestVerification;

const cookieExpiry = new Date(Date.parse(expires));

// Fetching Logic //

function MakeCalendarRequest(startDate, endDate){
  console.log("Fetching calendar from Nucleus starting: " + startDate + ", ending " + endDate); 
  //
  console.info("Requesting https://itv.dzjintonik.eu/CentralPlan/MyCentralPlan");
  let firstResponse = UrlFetchApp.fetch("https://itv.dzjintonik.eu/CentralPlan/MyCentralPlan", { headers: {
    "cookie": formatCookiesToString()
  }});
  let headers = firstResponse.getAllHeaders();
  console.log("Headers:");
  console.log(headers);
  let cookies = formatCookies(headers['Set-Cookie']);
  console.log("New cookies from server:");
  console.log(cookies);
  if (".AspNet.Cookies" in cookies){
    console.warn("New value for .AspNet.Cookies" + cookies[".AspNet.Cookies"])
    userProps.setProperty(".AspNet.Cookies", cookies[".AspNet.Cookies"]);
    throw "New .AspNet.Cookies value sent from server";
  }

  let sessionCookies = {};
  sessionCookies = {...sessionCookies, ...cookies}

  for(retrys = 0; retrys < 3; retrys++){
    console.info("Requesting calendar file, attempt " + (retrys + 1));
    let response = UrlFetchApp.fetch("https://itv.dzjintonik.eu/ResourceCalendar/ExportDataToFile?ResourceId="+resourceId+"&DateFrom="+startDate+"T00%3A00%3A00%2B01%3A00&DateTo="+endDate+"T00%3A00%3A00%2B01%3A00", {
      headers: {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
        "sec-ch-ua": "\"Not_A Brand\";v=\"99\", \"Google Chrome\";v=\"109\", \"Chromium\";v=\"109\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\"",
        "sec-fetch-dest": "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "same-origin",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "cookie": formatCookiesToString(sessionCookies),
        "Referer": "https://itv.dzjintonik.eu/CentralPlan/MyCentralPlan",
        "Referrer-Policy": "strict-origin-when-cross-origin"
      },
      method: "get"
    })
    headers = response.getAllHeaders();
    console.log("Response headers:");
    console.log(headers);
    if(headers['Set-Cookie']){
      cookies = formatCookies(headers['Set-Cookie']);
      console.log("New cookies from server:");
      console.log(cookies);
      sessionCookies = {...sessionCookies, ...cookies}
    }

    if (headers['Content-Type'].startsWith("text/calendar")){
        console.log("Calendar fetched successfully.")
        if(headers["Set-Cookie"]){
          console.info("Set-Cookie:" + headers["Set-Cookie"])
        }
        return response.getContentText();
      } else {
        console.log("Response code: " + response.getResponseCode());
        if(headers["Content-Type"].startsWith("text/html")){
          console.warn("Authentication failed");
        }
        //throw "Response is not calendar format";
      }
  }
  throw "Response is not calendar format";
}

function formatCookies(data){
  let output = {};

  if (typeof data == "string"){
    data = [data]
  }

  data.forEach(cookie => {
    let firstSplit = cookie.split(/=(.*)/s);
    let key = firstSplit[0];
    let secondSplit = firstSplit[1].split(/;(.*)/s);
    let value = secondSplit[0];
    //console.log(key, value);
    output[key] = value;
  })
  //console.log(output);
  return output;
}

function formatCookiesToString(sessionCookies){
  if (typeof sessionCookies != "object"){
    sessionCookies = "";
  }
  let extraCookies = ""

  for (const [key, value] of Object.entries(sessionCookies)){
      extraCookies = extraCookies + " " + key + "=" + value + ";";
    }

  let output = "PersonId="+personId+"; ProgramId=; MasterProgramId=null; InternalCompanyId=1; .AspNet.Cookies="+cookie+";" + extraCookies;
  return output;
}

// Calendar Logic //

function getOrCreateCalendar(){
  if (options.useDefaultCalendar) {
    console.log("fetching default calendar")
    return CalendarApp.getDefaultCalendar();
  } else {
    console.log("Looking for calendar named " + options.calendarName);
    let calendars = CalendarApp.getOwnedCalendarsByName(options.calendarName);
    console.log('Found %s matching calendars.', calendars.length)
    if (calendars.length == 1){
      return calendars[0];
    } else if (calendars.length > 1){
      throw "Found multiple calendars with configured name: " + options.calendarName
    } else if (calendars.length < 1){
      console.log("Calendar not found, creating new calendar named " + options.calendarName);
      let newCalendar = CalendarApp.createCalendar(options.calendarName, {color: CalendarApp.Color.LIME});
      return newCalendar;
    }
  }
}

// Returns array of CalendarEvents that start on a given date
function findEventsForDay(calendar, date){
  date.setHours(0);
  let events = calendar.getEventsForDay(date);
  let output = [];
  events.forEach(event => {
    let startTime = event.getStartTime();
      console.log("Start time: ", startTime)
      console.log("Date: ", date)
    if (!event.getTag("AutoEvent")){
      console.log("ignoring user created event");
    } else if (startTime < date){
      console.log("Ignoring event starting yesterday");
    } else {
      output.push(event);
    }    
  })
  console.log("Found %s existing auto events in calendar.", output.length)
  return output;
}

/**
 * @param {Array.Object} events Array of objects containing data for each calandar event
 * @param {String} startDate String representation of date the updating period starts on
 * @param {String} endDate String representation of date the updating period ends on
 */
function UpdateCalendar(events, startDate, endDate){
  console.log("Updating calendar for range: " + startDate + " - " + endDate);
  startDate = Date.parse(startDate);
  endDate = Date.parse(endDate);
  const oneDay = 8.64e7
  let calendar = getOrCreateCalendar();

  // For each day in range, check if there is an event starting on that date already in calendar, check if there is an event in the events array and update/remove accordingly.
  for(let i = startDate; i <= endDate; i += oneDay){
    let date = new Date(i);
    console.log("Updating: " + date);
    let existingEvents = findEventsForDay(calendar, date);
    let nucleusEventForDay = events.find((element) => {
      if(element.startTime.getDate() == date.getDate() && element.startTime.getMonth() == date.getMonth()){
        return true;
      }
      })
    if (nucleusEventForDay){
      console.log("Event found for day on Nucleus: " + nucleusEventForDay.title);
      UpdateDay(calendar, nucleusEventForDay, existingEvents);
    } else {
      console.log("No events for day on Nucleus.");
      existingEvents.forEach(event => {
        event.deleteEvent();
      })
    }
  }
}

function AddNewEvent(calendar, event){
  console.log("Writing new event to calendar");
  let newEvent;

  if (event.allDay){
    // Add one to end date as calendar API treats end date as exlusive for all day events
    event.endTime.setDate(event.endTime.getDate() + 1);
    newEvent = calendar.createAllDayEvent(event.title, event.startTime, event.endTime, event.options);
  } else {
    newEvent = calendar.createEvent(event.title, event.startTime, event.endTime, event.options);
  }
  newEvent.setTag("AutoEvent", "true");
  return newEvent;
}

function UpdateExistingEvent(calendar, oldEvent, newEvent){
    console.log("New event start time:", newEvent.startTime);
    console.log("New event end time:", newEvent.endTime);
    // If old event is not same type of event as new event, delete and re-add.
    if (oldEvent.isAllDayEvent() != newEvent.allDay){
      AddNewEvent(calendar, newEvent);
      oldEvent.deleteEvent();
      return;
    }

    if(oldEvent.getTitle() != newEvent.title){
        console.log("Updating event title to: " + newEvent.title)
        oldEvent.setTitle(newEvent.title);
      }
    // Check times match
    if ((oldEvent.getStartTime().getTime() != newEvent.startTime.getTime()) || (oldEvent.getEndTime().getTime() != newEvent.endTime.getTime())){
      // Skip time update if all day event
      if (!newEvent.allDay){
        console.log("Updating times to: " + newEvent.startTime + " to " + newEvent.endTime);
        oldEvent.setTime(newEvent.startTime, newEvent.endTime);
      }      
    }
    // Check description matches
    if(oldEvent.getDescription() != newEvent.options.description){
      console.log("Updating description.");
      oldEvent.setDescription(newEvent.options.description)
    }
    // Check location matches
    if (oldEvent.getLocation() != newEvent.options.location){
      console.log("Updating location to: " + newEvent.options.location)
      oldEvent.setLocation(newEvent.options.location);
    }
}

function UpdateDay(calendar, nucleusEvent, existingEvents){
  if (existingEvents.length < 1){
    // Add new event
    try {
      AddNewEvent(calendar, nucleusEvent);
    } catch (err){
      console.error(err)
    }    
  } else if (existingEvents.length == 1){
    // Check if event needs updating
    console.log("Checking whether existing event needs updating.")
    UpdateExistingEvent(calendar, existingEvents[0], nucleusEvent);
  } else {
    // Found multiple events to update
    throw "Found multiple matching events to update on " + date;
  }
}

function testUpdateCalendar(){
  UpdateCalendar(FormatCalendarEvents(testEvents), FormatDateForRequest(new Date(Date.now())), FormatDateForRequest(Date.now() + (12096e5*3)));
}

// Formatting and Parsing Logic //

// Returns date in format: 2005-12-31
const FormatDateForRequest = function(inputDate){
  let date = new Date(inputDate);
  let year = date.getFullYear();
  let month = date.getMonth() + 1;
  let day = date.getDate();

  if (month < 10){
    month = "0" + month
  }

  if (day < 10){
    day = "0" + day
  }
  return year + "-" + month + "-" + day
}

const FormatDateFromCalendar = function(input) {
  let output = input.substring(0,4)+"-"+input.substring(4,6)+"-"+input.substring(6,11)+":"+input.substring(11,13)+":"+input.substring(13);
  return output; 
}

const ParseDescription = function(input){
  let lines = input.split("\\n")
  let output = {};
  lines.forEach(line => {
    let thisLine = line.split(/: (.*)/s);
    let key = thisLine[0];
    let value = thisLine[1];

    switch (key){
      case "Booked Assets":
        value = value.match(/(?<=, |^)[^,]*(?=, Corrie)/g);
        break;
      case "Other Booked People":
        value = value.split(", ");
        break;
    }

    // Strip spaces out of keys
    key = key.replaceAll(" ", "");
    // Make first letter of key lowecase
    key = key.charAt(0).toLowerCase() + key.slice(1)

    output[key] = value;
  })

  return output;
}

const FormatDescriptionForEvent = function(input){
  let output = "";

  // Add role to event description
  output = output + "Role:\n"
  output = output + "  " + input.role + "\n"

  // Add location to event description
  if (input.bookedAssets != null){
      output = output + "Locations:\n";
      input.bookedAssets.forEach(location => {
      output = output + "  " + location + "\n";
    })
  }

  // Add crew list to event description
  output = output + "Crew:\n"
  input.otherBookedPeople.forEach(person => {
    output = output + "  " + person + "\n";
  })
  
  return output;
}

const ParseCalendar = function(inputString){
  let lines = inputString.split("\r\n");
  let events = [];
  let currentEvent = {};
  let active = false;
  lines.forEach(line => {
    if (line == 'BEGIN:VEVENT'){
      active = true;
      currentEvent = {};
      return;
    }
    if (line == 'END:VEVENT'){
      active = false;
      events.push(currentEvent);
      return;
    }
    if (active){
      let output = {}
      let thisLine = line.split(/:(.*)/s);
      let key = thisLine[0];
      let value = thisLine[1];

      switch (key){
        case "DTSTART":
          key = "start"
          value = new Date(FormatDateFromCalendar(value));
          output[key] = value;
          break;
        case "DTEND":
          key = "end"
          value = new Date(FormatDateFromCalendar(value));
          output[key] = value;
          break;
        case "SUMMARY":
          key = "title"
          value = value.replace(" - Coronation Street 2023", "")
          output[key] = value;
          break;
        case "DESCRIPTION":
          output = ParseDescription(value);
          break;
      }

      currentEvent = {...currentEvent, ...output}
    }
  });
  return events;
}

function FormatCalendarEvents (inputArray){
  // Produces array of event objects in format:
  // { title: Event Title,
  //   startTime: Event start,
  //   endTime: Event end,
  //   options: {
  //     description: Event description,
  //     location: Event location
  //   }
  // }
  let output = [];

  inputArray.forEach(event => {
    let thisEvent = {};
    let options = {};

    thisEvent.title = event.title;
    thisEvent.startTime = event.start;
    thisEvent.endTime = event.end;

    // Add one day to end time if end time is before start time (Some overnight events are supplied with incorrect end dates)
    if (thisEvent.endTime.getTime() < thisEvent.startTime.getTime()){
      thisEvent.endTime.setDate(thisEvent.endTime.getDate() + 1);
    }
    
    options.description = FormatDescriptionForEvent(event);
    //TODO: Locations not parsed currently
    options.location =  ""
    
    thisEvent.options = options;

    // Change formatting for all day events (annual leave, etc.)
    thisEvent.allDay = false;

    if (thisEvent.startTime.getHours() == 0 && thisEvent.endTime.getHours() == 23){
      thisEvent.allDay = true;
      thisEvent.title = event.role;
      thisEvent.options.description = "";
    }

    output.push(thisEvent);
  })

  return output;
}

function FetchAndUpdateCalendar(){
  // test user input
  if (!cookie || typeof cookie != "string") {
    throw "Cookie not found or invalid"
  }
  if (!expires || typeof expires != "string") {
    throw "Cookie 'expires' value not found or invalid"
  }
  if (!personId || typeof personId != "number") {
    throw "personId not found or invalid"
  }

  const fortnightsToFetch = 26;
  const failDays = 1;
  let startDate = FormatDateForRequest(Date.now());
  let endDate = FormatDateForRequest(Date.now() + (12096e5*fortnightsToFetch));
  // Fetch calendar from nucleus
  let calendarData;
  // Check if cookie has expired and throw error if it has
  if(cookieExpiry && cookieExpiry < Date.now()){
    throw "AspNet.Cookies Cookie Expired";
  }
  try{
    calendarData = MakeCalendarRequest(startDate, endDate);
    // Convert calendar into JS Object
    let events = ParseCalendar(calendarData);

    // Remove unused data and format for creating calendar events
    let calendarEvents = FormatCalendarEvents(events);

    // Update user calendar
    UpdateCalendar(calendarEvents, startDate, endDate);
    // Update user props to
    userProps.setProperty("lastSuccess", Date.now().toString());
  } catch (err) {
    let lastSuccess = userProps.getProperty("lastSuccess");
    console.log("lastSuccess from userProps:", lastSuccess);
    if (lastSuccess && +lastSuccess < (Date.now() - (8.64e+7 * failDays))){
      console.warn(err);
      throw "Script has not completed within " + failDays + " day(s)";
    } else {
      console.log("Script has failed to complete but has completed successfully within the last " + failDays + " day(s)");
      console.warn(err);
    }
  }
}
