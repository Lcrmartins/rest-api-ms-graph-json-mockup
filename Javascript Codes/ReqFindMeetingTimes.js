const DateTimeMaker = (date, time) => {
    return date + 'T' + time + ':00Z';
};

function execute(context, proxy) {
    const accessToken = context.GetVariable('token');
    const roomName = context.GetVariable('room_name');
    const roomEmail = context.GetVariable('room_email');
    const startDate = context.GetVariable('start_date');    // must be in (YYYY-MM-DD) format
    const endDate = context.GetVariable('end_date');        // must be in (YYYY-MM-DD) format
    const startTime = context.GetVariable('start_time');    // must be in (HH:MM:SS) format and in UTC
    const endTime = context.GetVariable('end_time');        // must be in (HH:MM:SS) format and in UTC
    const duration = 'PT'+context.GetVariable('duration');  // must be in (hhHmmM) format (like 1H30M -> PT1H30M ou 45M -> PT45M )

    const startDateTime = DateTimeMaker(startDate, startTime);
    const endDateTime = DateTimeMaker(endDate, endTime);

    const url = 'https://graph.microsoft.com/v1.0/me/findMeetingTimes';
    const method = 'POST';
    const contentType = 'application/json';
    const headers = {
        "Authorization": accessToken
    };

    const jsonBody = `{
        "attendees": [
            {
                "emailAddress": {
                    "address": ${roomEmail},
                    "name": ${roomName}
                },
                "type": "Required"
            }
        ],
        "timeConstraint": {
            "timeslots": [
                {
                    "start": {
                        "dateTime": ${startDateTime},
                        "timeZone": "Pacific Standard Time"
                    },
                    "end": {
                        "dateTime": ${endDateTime},
                        "timeZone": "Pacific Standard Time"
                    }
                }
            ]
        },
        "meetingDuration": ${duration}
    }`;
           
    try
    {
        var response = Request({ Url: url, Method: method, Body: jsonBody, ContentType: contentType, Headers:headers });
        
        
        if (response.IsSuccessStatusCode) {
            
            if (response.Result.emptySuggestionsReason) {
                return response.Result.emptySuggestionsReason;
            }
            let suggestions = '';
            let suggestion = '';
            for (let i = 0; i < response.Result.meetingTimeSuggestions.length; i++){
                let date = response.Result.meetingTimeSuggestions[i].meetingTimeSlot.start.dateTime.substring(0, 9);
                let from = response.Result.meetingTimeSuggestions[i].meetingTimeSlot.start.dateTime.substring(11, 15);
                let to = response.Result.meetingTimeSuggestions[i].meetingTimeSlot.end.dateTime.substring(11, 15);
                suggestion = `[c:link label=${date} from ${from} to ${to} value= Meeting Suggestion ${i}]`;   
                suggestions += suggestion;
            }

            return suggestions;
        }
        return false;
    } catch(err) {
        return false;
    }
}



