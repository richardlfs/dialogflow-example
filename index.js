/**
 * Copyright 2017 Google Inc. All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

'use strict';

const functions = require('firebase-functions');
const {google} = require('googleapis');
const {WebhookClient} = require('dialogflow-fulfillment');

// Enter your calendar ID below and service account JSON below
const calendarId = "xxxxxxxxxxxxxxxxxxxx@group.calendar.google.com"; //這邊換成你的calendar ID

const serviceAccount = {
 "type": "service_account",
 "project_id": "",
 "private_key_id": "",
 "private_key": "",
 "client_email": "",
 "client_id": "",
 "auth_uri": "",
 "token_uri": "",
 "auth_provider_x509_cert_url": "",
 "client_x509_cert_url": "",
 "universe_domain": ""
}
; // 這邊換成你自己的JSON金鑰 從 {"type": "service_account",...開始複製貼上全部


// Set up Google Calendar Service account credentials
const serviceAccountAuth = new google.auth.JWT({
  email: serviceAccount.client_email,
  key: serviceAccount.private_key,
  scopes: 'https://www.googleapis.com/auth/calendar'
});

const calendar = google.calendar('v3');
process.env.DEBUG = 'dialogflow:*'; // enables lib debugging statements

const timeZone = 'Asia/Hong_Kong'; // 手動設定時區

exports.dialogflowFirebaseFulfillment = functions.https.onRequest((request, response) => {
  const agent = new WebhookClient({ request, response });
  console.log("Parameters", agent.parameters);
  const leave_type = agent.parameters.leave_type; //leave_type 是你在DialogFlow設定BOT會讀取到的Entity
  const person = agent.parameters.person.name;  //person 也是你在DialogFlow設定BOT會讀取到的Entity(person是系統內建的)
  
  function makeAppointment (agent) {
    // Calculate appointment start and end datetimes (end = +1hr from start)
    //console.log("Parameters", agent.parameters.date);
    const dateTimeStart = new Date(agent.parameters['date-time']); //date-time' 也是你在DialogFlow設定BOT會讀取到的Entity(date-time是系統內建的)
    const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));
    const appointmentTimeString = dateTimeStart.toLocaleString(
      'en-US',
      { month: 'long', day: 'numeric', hour: 'numeric', timeZone: timeZone }
    );

    // Check the availibility of the time, and make an appointment if there is time on the calendar
    return createCalendarEvent(dateTimeStart, dateTimeEnd, leave_type, person).then(() => {
      agent.add(`好的，您的請假資訊已登記 ${appointmentTimeString} !.`); 
    }).catch(() => {
      agent.add(`I'm sorry, there are no slots available for ${appointmentTimeString}.`);
    });
  }

  let intentMap = new Map();
  intentMap.set('申請請假與填報資訊', makeAppointment); //這裡要填你在DialogFlow設定的Intent
  agent.handleRequest(intentMap);
});



function createCalendarEvent (dateTimeStart, dateTimeEnd, leave_type, person) {
  return new Promise((resolve, reject) => {
    calendar.events.list({
      auth: serviceAccountAuth, // List events for time period
      calendarId: calendarId,
      timeMin: dateTimeStart.toISOString(),
      timeMax: dateTimeEnd.toISOString()
    }, (err, calendarResponse) => {
      // Check if there is a event already on the Calendar
      if (err || calendarResponse.data.items.length > 0) {
        reject(err || new Error('Requested time conflicts with another appointment'));
      } else {
        // Create event for the requested time period
        calendar.events.insert({ auth: serviceAccountAuth,
          calendarId: calendarId,
          resource: {summary: person + leave_type, description: person +'假別'+ leave_type,
            start: {dateTime: dateTimeStart},
            end: {dateTime: dateTimeEnd}}
        }, (err, event) => {
          err ? reject(err) : resolve(event);
        }
        );
      }
    });
  });
}


