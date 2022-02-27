const getToken = (context) => {
    let apiURL = context.GetVariable('office365_auth_api_url');
    var payload = `client_id=${context.GetVariable("ad_office365_client_id")}&scope=${context.GetVariable("ad_office365_scope")}&client_secret=${context.GetVariable("ad_office365_client_secret")}&grant_type=${context.GetVariable("ad_office365_grant_type")}`;
    try
    {
        var response = Request({ Url: apiURL, Method: 'POST', Body: payload, ContentType: "application/x-www-form-urlencoded"});
        if(response.IsSuccessStatusCode) {
            context.UpdateInputVariable("office365_api_access_token", response.Result.access_token);
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