// postman發req時的大致內容
import fetch from "node-fetch";
var myHeaders = new Headers();
myHeaders.append(
  "Cookie", "CMP=o6rf4tadr8f2shcfb0kvlqs3bs; CMPUT=1675495892; CMPEP=1675500721.9d9a4fa00e68b8b; TS01665764=0107dddfefba65f1e309db85986b7d51c65018deabc99003325d6796c3a35a5edfdfc34386959ea3ebe51bdb19e204141258d20ff92005c32236dfc980ea76f0caee3cb56a92b12b16b7562c4281e32fe373fd29dd907010f702e359a0bf3c4f37f24f313c9be13c57823ae362fe66e0c92ec54c96951409da2b769aafac28251fddf90e81278ca803d38d9621926b6c3d2577f20c2028ef19592f882a74ae8a170bce9c9f"
  );

var formdata = new FormData();
formdata.append("type", "table_year");
formdata.append("stn_ID", "C0X170");
formdata.append("stn_type", "auto_C0");
formdata.append("start", "2022-01-01T00:00:00");
formdata.append("end", "2022-12-31T00:00:00");

var requestOptions = {
  method: 'POST',
  headers: myHeaders,
  body: formdata,
  redirect: 'follow'
};

fetch("https://codis.cwb.gov.tw/api/station", requestOptions)
  .then(response => response.text())
  .then(result => console.log(result))
  .catch(error => console.log('error', error));