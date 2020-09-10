
/** * @OnlyCurrentDoc */

function onOpen () {
  SpreadsheetApp.getUi()
  .createMenu(
    'Calendar Tool'
    )
  .addItem(
    'Activate (run me first)','activate'
    )
  .addItem(
    'Fetch classes','fetchClasses'
    )
  .addItem(
    'Create Events','createEvents'
    )  
  .addToUi()
}

function activate () {
  SpreadsheetApp.getUi().alert('You should be ready to go now! Hopefully you clicked through all the permissions');
}

function fetchClasses () {
  let coursesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Classes')
  coursesSheet.clear();
  coursesSheet.appendRow(['Course ID','Name','Section','Calendar ID','Group Email','Enrollment Code']);
  let courses = Classroom.Courses.list({teacherId:'me',
                                        courseStates:['ACTIVE','PROVISIONED']}).courses.map(
    (course)=>[course.id,course.name,course.section,course.calendarId,course.courseGroupEmail,course.enrollmentCode]
    );
  Logger.log(courses);
  courses.map((r)=>coursesSheet.appendRow(r));
}


function createEvents () {
  const meetingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings')
  let table = SHL.Table(meetingSheet.getDataRange())
  table.forEach(
    (row)=>{
    if (!row.Added && row.Start) { // not the header row...
       let series = CalendarApp.getCalendarById(row.Calendar)
       .createEventSeries(
         row.Title,
         toDate(row.Start),
         toDate(row.End),
         CalendarApp.newRecurrence()
         .addWeeklyRule()
         .until(toDate(row.Until))         
         );
      row.Added = series.getId();
    if (row.Invite) {
      series.addGuest(row.Invite);
    }
    }
});
}

function toDate (timeString) {
  if (timeString.getDate) {return timeString} // timeString already a date?
  timeString = timeString.replace(/-/g,"/");
  var d = new Date(timeString);
  Logger.log('toDate %s=>%s',timeString,d);
  return d;
}