const logMeIn = (context) => {
    const TENANT_ID = 'fa453bd6-98a2-45ad-bcbb-c4fdb33edacc';
    const APP_ID = '5cf3fb1a-8543-45f9-a210-bdf3879f243c';
    const APP_SECERET = 'xC47Q~o3CiovBddMOzTnEThzSLQruJor~vc2I';
    const TOKEN_ENDPOINT ='https://login.microsoftonline.com/'+TENANT_ID+'/oauth2/v2.0/authorize?';
    const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
    const REDIRECT_URI = 'https%3A%2F%2Fassistant.demo.sophie.chat%2FLuis';
    const GRANT = 'client_credentials';
    const RESPONSE_TYPE = 'code';
    const RESPONSE_MODE = 'query';
    const SCOPE = 'offline_access%20user.read%20calendars.readwrite.shared';

    const MYSTATE = 'MyState321456';

    const client_idPARAM = 'client_id=' + APP_ID;
    const response_typePARAM = '&response_type=' + RESPONSE_TYPE;
    const redirect_uriPARAM = '&redirect_uri=' + REDIRECT_URI;
    const response_modePARAM = '&response_mode=' + RESPONSE_MODE;
    const scopePARAM = '&scope=' + SCOPE;
    const statePARAM = '&state=' + MYSTATE;
    const params = client_idPARAM + response_typePARAM + redirect_uriPARAM + response_modePARAM + scopePARAM + statePARAM;

    const Url = TOKEN_ENDPOINT + params;
    try {
        var response = Request({ Url: Url });
        if (response.IsSuccessStatusCode) {
            
            return response.Result;
        }
        return false;
    } catch (err) {
        return false;
    }
};
function execute(context, proxy) {
    return logMeIn(context);
}

