const TENANT_ID = 'fa453bd6-98a2-45ad-bcbb-c4fdb33edacc';
const APP_ID = '5cf3fb1a-8543-45f9-a210-bdf3879f243c';
const APP_SECERET = 'xC47Q~o3CiovBddMOzTnEThzSLQruJor~vc2I';
const TOKEN_ENDPOINT ='https://login.microsoftonline.com/'+TENANT_ID+'/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const GRANT = 'client_credentials';
const axios = require('axios');
const qs = require('qs');

const postData = {
  client_id: APP_ID,
  scope: MS_GRAPH_SCOPE,
  client_secret: APP_SECERET,
  grant_type: GRANT
};

axios.defaults.headers.post['Content-Type'] =
  'application/x-www-form-urlencoded';

axios
  .post(TOKEN_ENDPOINT, qs.stringify(postData))
  .then(response => {
    console.log(response.data);
  })
  .catch(error => {
    console.log(error);
  });