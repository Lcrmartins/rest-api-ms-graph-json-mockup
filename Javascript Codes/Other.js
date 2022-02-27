const buildMailNickname = (displayName) => {
    let splittedDisplayName = displayName.includes(",") ? displayName.replace(" ", "").split(",") : displayName.split(" ");
    return `${splittedDisplayName[0]}${splittedDisplayName[splittedDisplayName.length - 1][0]}`;
};
function execute(context, proxy) {
    let accessToken = proxy.ExecuteDynamicIntegration('get_ad_office365_token', context);
    if(accessToken === "false")
    {
        return "[c:newline][s:redirect rule=office365_mgm_general_error][c:newline]";
    }
    let headers = {
        "Authorization": accessToken,
    };
    context.updateInputVariable('password_length', 16);
    let password = proxy.ExecuteDynamicIntegration('password_generator', context);
    const displayName = context.GetVariable("office365_user_display_name");
    const mailNickname = buildMailNickname(displayName);
    const newUsername = context.GetVariable("office365_username").trim();
    const emailDomain = context.GetVariable("office365_mgm_email_domain");
    let payload = {
      "accountEnabled": "true",
      "displayName": displayName,
      "mailNickname": mailNickname,
      "userPrincipalName": `${newUsername}${emailDomain}`,
      "usageLocation": "US",
      "passwordProfile": {
        "forceChangePasswordNextSignIn": "false",
        "password": password
      }
    };
    let apiURL = context.GetVariable('office365_mgm_api_url');
    try
    {
        var response = Request({ Url: `${apiURL}/users`, Method: 'POST', Body: payload, Headers: headers});
        context.UpdateInputVariable("created_user_response", JSON.stringify(response));
        if(response.IsSuccessStatusCode){
            context.UpdateInputVariable("office365_user_id", response.Result.id);
            context.UpdateInputVariable("office365_user_email", response.Result.userPrincipalName);
            context.UpdateInputVariable("office365_user_password", password);
            return `[s:redirect rule=office365_user_created]`;
        }
        return `[s:redirect rule=office365_failed_to_create_user]`;
    } catch(err) {
        context.UpdateInputVariable("err", err);
        return `[s:redirect rule=office365_mgm_general_error]`;
    }
}