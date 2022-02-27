const getToken = (context) => {
    const TENANT_ID = 'fa453bd6-98a2-45ad-bcbb-c4fdb33edacc';
    const APP_ID = '5cf3fb1a-8543-45f9-a210-bdf3879f243c';
    const APP_SECRET = 'xC47Q~o3CiovBddMOzTnEThzSLQruJor~vc2I';
    const TOKEN_URL ='https://login.microsoftonline.com/'+TENANT_ID+'/oauth2/v2.0/token';
    const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
    const GRANT = 'client_credentials';
    var payload = `client_id=${APP_ID}&scope=${MS_GRAPH_SCOPE}&client_secret=${APP_SECRET}&grant_type=${GRANT}`;
    try
    {
        var response = Request({ Url: TOKEN_URL, Method: 'POST', Body: payload, ContentType: "application/x-www-form-urlencoded"});
        if(response.IsSuccessStatusCode) {
            context.UpdateInputVariable("calendar_access_token", response.Result.access_token);
            return response.Result.access_token;
        }
        return false;
    } catch(err) {
        return false;
    }
};
function execute(context, proxy) {
    return getToken(context);
}