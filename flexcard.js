var today // ตัวแปรวันนี้
var thisMonth //ตัวแปรรับค่าเดือน
var ssId = "1HL7rQFPJHbAyN_emtwWrKugYlbgbwLk6DheZL46cd9U";  // gg Sheet ID
var ss = SpreadsheetApp.openById(ssId);
var sheet = ss.getSheetByName("Form");  // ชื่อหน้าใน ggsheet
var token = "Bearer MVfD+dKoFGs2yzXAp2uURnxnAaRc+UTcH+zmulgajWmMsHHuagd1lRmXj/Vke/0zru5qeCcCGQVSMMtjLHibxEIOuJIkVTnKbjp2TUpCyyMobR03mqPG6a4R4AGcRvgNvnuocy7yAciO0umDpP2mhQdB04t89/1O/w1cDnyilFU=" //access token
// var roomId = "C54568fbcca99fd7e8e4240d0540f8fce" // เลขห้อง 
var roomId = "Cdbd53f366b99fc8eaa47385d7c12297f" // เลขห้อง


function getMonth() {
  const month = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฏาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  today = new Date(); // เอาไว้ set 7 วัน 15 วัน
  thisMonth = month[today.getMonth()]; //ใช้ ตรวจเดือน
  return thisMonth
}




function getReport() {
  const header = getMonth();
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headers = values.shift();
  const columnIndex = headers.indexOf(header);
  const data = values

  //msg.messages[0].contents.body.contents.push(obj)
  var msg = {
    "to": roomId,
    "messages": [
      {
        "type": "flex",
        "altText": "ติดตาม KPI",
        "contents": {
          "type": "bubble",
          "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "action": {
              "type": "uri",
              "label": "Action",
              "uri": "https://linecorp.com"
            },
            "contents": [
              {
                "type": "text",
                "text": "Bot Report | QDFM",
                "size": "xs",
                "color": "#00FF08FF",
                "contents": []
              },
              {
                "type": "box",
                "layout": "vertical",
                "contents": [
                  {
                    "type": "text",
                    "text": "ติดตามการทำงาน",
                    "weight": "bold",
                    "size": "sm",
                    "color": "#7C7C7CFF",
                    "contents": []
                  },
                  {
                    "type": "text",
                    "text": "แจ้งเตือนล่วงหน้า 3 วัน",
                    "weight": "bold",
                    "size": "xl",
                    "contents": []
                  }
                ]
              },
              {
                "type": "text",
                "text": "ฝ่ายวิศวกรรมบริการและอาคารสถานที่",
                "size": "sm",
                "color": "#AAAAAA",
                "align": "start",
                "wrap": false,
                "contents": []
              },
              {
                "type": "text",
                "text": "วันศุกร์ที่ 29 เมษายน พ.ศ.2565",
                "size": "sm",
                "color": "#000000FF",
                "wrap": true,
                "contents": []
              },
              {
                "type": "box",
                "layout": "vertical",
                "contents": [
                  {
                    "type": "spacer"
                  }
                ]
              },

            ]
          }
        }

      }
    ]
  }


  for (i = 0; i < data.length; i++) {

    let row = data[i]

    // Logger.log(row[0] + row[columnIndex])
    var obj = {
      "type": "box",
      "layout": "horizontal",
      "contents": [
        {
          "type": "text",
          "text": row[0],
          "size": "sm",
          "color": "#727272FF",
          "flex": 2,
          "align": "start",
          "contents": []
        },
        {
          "type": "text",
          "text": row[columnIndex],
          "size": "sm",
          "color": "#C9C9C9FF",
          "align": "end",
          "gravity": "center",
          "contents": []
        }
      ]
    }
    msg.messages[0].contents.body.contents.push(obj)

  } //end loop for
  msg.messages[0].contents.body.contents.push({
    "type": "separator"
  })
  msg.messages[0].contents.body.contents.push({
    "type": "box",
    "layout": "horizontal",
    "contents": [
      {
        "type": "text",
        "text": "รวมทั้งหมด",
        "weight": "bold",
        "flex": 2,
        "align": "start",
        "contents": []
      },
      {
        "type": "text",
        "text": "10 ทีม",
        "weight": "bold",
        "size": "xl",
        "color": "#38C21AFF",
        "align": "end",
        "contents": []
      }
    ]
  })
  SendLine(msg, token)
}






function SendLine(data, token) {
  var url = "https://api.line.me/v2/bot/message/push";
  var lineHeader = {
    "Content-Type": "application/json",
    "Authorization": token
  };

  var postData = data

  var options = {
    "method": "POST",
    "headers": lineHeader,
    "payload": JSON.stringify(postData)
  };



  try {
    var response = UrlFetchApp.fetch(url, options);
  }

  catch (error) {
    Logger.log(error.name + "：" + error.message);
    return;
  }

  if (response.getResponseCode() === 200) {
    Logger.log("Sending message completed.");
  }
}





