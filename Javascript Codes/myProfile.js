const getMyProfile = (context) => {
    const PROFILE_URL = 'https://graph.microsoft.com/v1.0/me/';
    const accessToken = context.GetVariable("calendar_access_token");
    let headers = {
        "Authorization": accessToken
    };
    try
    {
        var response = Request({ Url: PROFILE_URL, Method: 'GET', Headers: headers});
        if(response.IsSuccessStatusCode) {
            context.UpdateInputVariable("userDisplayName", response.Result.displayName);
            context.UpdateInputVariable("userMail", response.Result.mail);
            context.UpdateInputVariable("userId", response.Result.id);
            return response.Result;
        }
        return response.StatusCode+' \\n \\n'+response.ErrorResult;
    } catch(err) {
        return false;
    }
};
function execute(context, proxy) {
    return getMyProfile(context);
}