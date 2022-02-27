const getToken = (context) => {
    const TENANT_ID = 'fa453bd6-98a2-45ad-bcbb-c4fdb33edacc';
    const APP_ID = '5cf3fb1a-8543-45f9-a210-bdf3879f243c';
    const APP_SECRET = 'xC47Q~o3CiovBddMOzTnEThzSLQruJor~vc2I';
    const ENCODED_REDIRECT_URI = 'https%3A%2F%2Fassistant.demo.sophie.chat%2FLuis'; 
    const ENCODED_SCOPE = 'offline_access%20user.read%20calendars.readwrite.shared';
    const STATE = 'MyState321456';
    const TOKEN_URL =
        'https://login.microsoftonline.com/' +
        TENANT_ID +
        '/oauth2/v2.0/authorize?' +
        'client_id=' +
        APP_ID +
        '&response_type=code' +
        '&redirect_uri=' +
        ENCODED_REDIRECT_URI +
        '&response_mode=form_post' +
        '&scope=' +
        ENCODED_SCOPE +
        '&state=' +
        STATE;   
    const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
    const GRANT = 'client_credentials';
    var payload = `client_id=${APP_ID}&scope=${MS_GRAPH_SCOPE}&client_secret=${APP_SECRET}&grant_type=${GRANT}`;
    try
    {
        var response = Request({ Url: TOKEN_URL, Method: 'GET' });
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

https://login.microsoftonline.com/fa453bd6-98a2-45ad-bcbb-c4fdb33edacc/oauth2/v2.0/authorize?client_id=5cf3fb1a-8543-45f9-a210-bdf3879f243c&response_type=code&redirect_uri=https%3A%2F%2Fassistant.demo.sophie.chat%2FLuis&response_mode=query&scope=offline_access%20user.read%20calendars.readwrite.shared&state=MyState321456

let code = '0.AUYA1jtF-qKYrUW8u8T9sz7azBr781xDhflFohC984efJDyAACE.AQABAAIAAAD--DLA3VO7QrddgJg7Wevr86ziUL1vDgnb7E_wrcwNUi-sUWilzTaDFm0Y_WqEIcfWDDnkEStBkdy4qtrDGwRei59YAh-cWtuHcfDKIl8FGsYmleCD9PzkAry58yDIzRKWiQH8SxMZdqUkTDJZ-AIs5kKH-AClXw-QZkT6n_VNZjvtNCzxeBC9yJPT1Wmclw5QXcqk_ncktxT84-KkBSOOwOswznhPwqTtpiPPIdbhkY2ICQTv3AjrjSJ7Y7Yy1L6SsRp-AvCM-5gHOnSD5bMwaJlLRHTU6Hk_qLsmsqld_srH7EF5SGRezLiq_WkXIKpojEQ8gfGaK7gptsFgUX1wkXjP0wbxJEEKJYvJggZpoBM_H-2X0XTQyeFGAHHQrESj_UrXLLRO7W-uIMGk7nRurK4emw2kGGRq2brNwtsA93SVXMKxCzKXqV912DvN7ubJ48LmHRlyEFsAS0l_1urv96rUrCLfuYpmZe7BEj_rCADnIbDwtsLXok68xEY7IpCscOpgO4dKxTokmmLys4zgb0t0DEirIAnWFTglHaoaMuKqnRgKsehDKfGoKFPbugh1kF5FDiIoDfVUNofu6RXgTpCWK0aQou77XvumeVS6tWbSnrt3et8HYHcCBEk_LSQNsSR-PuVcGj1rrl_1SESP81NPuiH_BFI___byARVkA9e39UqSt2znUw0W8nB9d-hnewsl7BjQ0x_YWMdKmL_GbYnnG2_ZTnFsVMRtVtm3UiAA';
    
let Session_State = '&session_state=95276400-580e-4f7f-a622-011a07ee927b#';
    
let MyStateRes = 'MyState321456';
