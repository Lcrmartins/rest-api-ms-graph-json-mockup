const DateTimeMaker = (date, time) => {
    return date + 'T' + time + ':00Z';
};

// Obtaining the empty slots ...

const SlotFinder = (accessToken, idOrUserPrincipalName, roomEmail, roomName, startDateTime, endDateTime, durationCode, maxResultLength) => {
    const findMeetingTimesUrl = `https://graph.microsoft.com/users/${idOrUserPrincipalName}/findMeetingTimes`;
    const method = 'POST';
    const contentType = 'application/json';
    const headers = { "Authorization": accessToken };
    let payload = {
        "attendees": [
            {
                "emailAddress": {
                    "address": roomEmail,
                    "name": roomName
                },
                "type": "Required"
            }
        ],
        "timeConstraint": {
            "timeslots": [
                {
                    "start": {
                        "dateTime": startDateTime,
                        "timeZone": "E. South America Standard Time"
                    },
                    "end": {
                        "dateTime": endDateTime,
                        "timeZone": "E. South America Standard Time"
                    }
                }
            ]
        },
        "meetingDuration": durationCode
    };
    try {
        var response = Request({ Url: findMeetingTimesUrl, Method: method, Body: payload, ContentType: contentType, Headers:headers });
        if (response.IsSuccessStatusCode) {
            if (response.Result.emptySuggestionsReason) {
                return response.Result.emptySuggestionsReason;
            }
            let suggestions = '';
            let suggestion = '';
            for (let i = 0; i < Math.min(response.Result.meetingTimeSuggestions.length, maxResultLength) ; i++) {
                let date = response.Result.meetingTimeSuggestions[i].meetingTimeSlot.start.dateTime.substring(0, 9);
                let from = response.Result.meetingTimeSuggestions[i].meetingTimeSlot.start.dateTime.substring(11, 15);
                let to = response.Result.meetingTimeSuggestions[i].meetingTimeSlot.end.dateTime.substring(11, 15);
                suggestion = `[c:link label=${date} from ${from} to ${to} value= Meeting Suggestion ${i} for ${roomName}]`;   
                suggestions += suggestion;
            }
            return suggestions;
        }
        return false;
    } catch(err) {
        return false;
    }
};

// Main

function execute(context, proxy) {
    const roomOption = context.GetVariable('room_option');
    const startDate = context.GetVariable('start_date');            // must be in (YYYY-MM-DD) format
    const endDate = context.GetVariable('end_date');                // must be in (YYYY-MM-DD) format
    const startTime = context.GetVariable('start_time');            // must be in (HH:MM:SS) format and in UTC
    const endTime = context.GetVariable('end_time');                // must be in (HH:MM:SS) format and in UTC
    const durationCode = 'PT'+context.GetVariable('duration');      // must be in (hhHmmM) format (like 1H30M -> PT1H30M ou 45M -> PT45M )
    const room = [
        {
            "room 1": {
                "email": "room1@seethisapp.onmicrosoft.com",
                "name": "Room 1"
            }
        },
        {
            "room 2": {
                "email": "room2@seethisapp.onmicrosoft.com",
                "name": "Room 2"
            }
        },
        {
            "room 3": {
                "email": "room3@seethisapp.onmicrosoft.com",
                "name": "Room 3"
            }
        }
    ];

    const startDateTime = DateTimeMaker(startDate, startTime);
    const endDateTime = DateTimeMaker(endDate, endTime);

    // Obtaining the id Or UserPrincipalName and token...

    let idOrUserPrincipalName = proxy.ExecuteDynamicIntegration('Id_or_User_Name', context); // returns my email

    let accessToken = proxy.ExecuteDynamicIntegration('Calendar_Access_Token', context);     // returns a token 
    if(accessToken === "false")
    {
        return "[c:newline][s:redirect rule=graph_calendar_general_error][c:newline]";
    }
    let headers = {
        "Authorization": accessToken,
    };

    // Room options Switch
    var roomName = '';
    var roomEmail = '';
    var maxResultLength = 5;
    
    switch (roomOption){
        case "room 1": {
            roomName = room[0].name;
            roomEmail = room[0].email;
            return SlotFinder(accessToken, idOrUserPrincipalName, roomEmail, roomName, startDateTime, endDateTime, durationCode, maxResultLength);
        }
        case "room 2": {
            roomName = room[1].name;
            roomEmail = room[1].email;
            return SlotFinder(accessToken, idOrUserPrincipalName, roomEmail, roomName, startDateTime, endDateTime, durationCode, maxResultLength);
        }
        case "room 3": {
            roomName = room[2].name;
            roomEmail = room[2].email;
            return SlotFinder(accessToken, idOrUserPrincipalName, roomEmail, roomName, startDateTime, endDateTime, durationCode, maxResultLength);
        }
        case "No Preference": {
            maxResultLength = 2;
            var suggestions='';
            var suggestion='';
            for (let i = 0; i < room.length; i++){
                roomName = room[0].name;
                roomEmail = room[0].email;
                suggestion = SlotFinder(accessToken, idOrUserPrincipalName, roomEmail, roomName, startDateTime, endDateTime, durationCode, maxResultLength);
                suggestions+=suggestion;
            }
            return suggestions;
        }
    }
}