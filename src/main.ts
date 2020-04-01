const SLACK_VERIFICATIONTOKEN: string = PropertiesService.getScriptProperties().getProperty('SLACK_VERIFICATIONTOKEN');
const SLACK_WEBHOOK_URL: string = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
const SPREADSHEET_ID: string = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const SHEET1_NAME: string = PropertiesService.getScriptProperties().getProperty('SHEET1_NAME');
const SENDNOTFOUNDCOMMENTSTR: string = PropertiesService.getScriptProperties().getProperty('SENDNOTFOUNDCOMMENTSTR');
const SPREADSHEET_URL: string = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');

function doPost(e: string) {
    let verificationToken: string = e.parameter.token;
    if (verificationToken != SLACK_VERIFICATIONTOKEN) {
        throw new Error('Invalid Token');
    }
    
    let arg: string = e.parameter.text.replace(/　/g, ' ').trim();
    
    if (arg.length > 0) {
        let tmpAry: string[] = arg.split(' ');
        if (tmpAry[0] === 'URL') {
            PostMessageToSlack(SPREADSHEET_URL);
            return ContentService.createTextOutput();
        }
    }

    let trgtSpreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let trgtSh = trgtSpreadSheet.getSheetByName(SHEET1_NAME);

    let dataLastRow = trgtSh.getLastRow();
    let trgtRng = trgtSh.getRange(1, 1, dataLastRow, 3);
    let trgtAry: any[] = trgtRng.getValues();

    // 配列のインデックスは「0」から始まるため「-1」
    const phraseColIndex: number = 1 - 1;
    const descriptionColIndex: number = 2 - 1;
    const repeatFlgColIndex: number = 3 - 1;
    
    let flgRng = trgtSh.getRange(1,3, dataLastRow,3);
    let flgAry: number[][] = flgRng.getValues();
    let flgCnt: number = 0;
    flgAry.forEach(function(value){
        if (value[0] === 1) {
            flgCnt += value[0];
        }
    });
    
    // フラグのない行を取得
    let noFlgTrgtAry: any[] = new Array;
    const flgNum: number = 1;

    let trgtRowIndex: number;
    // 1周していたらフラグをリセット
    if (flgCnt === dataLastRow) {
        for (let i = 0; i <= dataLastRow - 1; i++) {
            trgtAry[i][repeatFlgColIndex] = 0;
        }
        noFlgTrgtAry = trgtAry
        trgtRowIndex = Math.floor(Math.random() * noFlgTrgtAry.length);

    }　else {
        trgtAry.forEach(function(el, index) {
            if (el[repeatFlgColIndex] != flgNum) {
                noFlgTrgtAry.push([el[phraseColIndex], el[descriptionColIndex], index]);
            }
        });
        trgtRowIndex = Math.floor(Math.random() * noFlgTrgtAry.length);
        const trgtIndexColIndex: number = 2;
        trgtRowIndex = noFlgTrgtAry[trgtRowIndex][trgtIndexColIndex];
    }
    
    let noMatchFlg: boolean = false;
    if (arg.length > 0) {
        let wordAry: string[] = arg.split(' ');
        let wordCnt: number = wordAry.length;

        let trgtWord: string;
        let trgtPhraseRowsIndexAry: number[] = new Array;
        
        let missFlg: boolean = false;
        trgtAry.forEach(function(el, index) {
            for (let i = 0; i < wordCnt; i++) {
                trgtWord = wordAry[i];

                if (el[phraseColIndex].toString().indexOf(trgtWord) === -1) {
                    missFlg = true;
                    break;
                }
            }
            if (missFlg === false) {
                trgtPhraseRowsIndexAry.push(index);
            }
            missFlg = false;
        });
        if (trgtPhraseRowsIndexAry.length === 0) {
            noMatchFlg = true;
        } else {
            let trgtIndex: number = Math.floor(Math.random() * trgtPhraseRowsIndexAry.length);
            trgtRowIndex = trgtPhraseRowsIndexAry[trgtIndex];            
        }
    
    } else {
        // 返したフレーズ行にフラグをたてる
        trgtAry[trgtRowIndex][repeatFlgColIndex] = flgNum;
        // フラグをシートに反映
        trgtRng.setValues(trgtAry);
    }
    
    let trgtPhrase: string = trgtAry[trgtRowIndex][phraseColIndex];
    let trgtDescription: string = trgtAry[trgtRowIndex][descriptionColIndex];

    let sendComment: string;
    if (noMatchFlg === true) {
        sendComment = `${ SENDNOTFOUNDCOMMENTSTR }`
    } else {
        sendComment = `\`${ trgtPhrase }\`

${ trgtDescription }`
    }
    
    PostMessageToSlack(sendComment);
    
    return ContentService.createTextOutput();
}

function PostMessageToSlack(sendBody: string) {
    let params: any = {
        method: 'post',
        contentType: 'application/json',
        payload: `{"text":"${ sendBody }"}`
    };
    
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, params);
}

// --------Bot用関数--------
const TARGETHOUR: string = PropertiesService.getScriptProperties().getProperty('TARGETHOUR');
const TARGETMINITUE: string = PropertiesService.getScriptProperties().getProperty('TARGETMINITUE');
const TRIGGERTARGETFUNCTION: string = PropertiesService.getScriptProperties().getProperty('TRIGGERTARGETFUNCTION');

function SendPhraseByBot {
    // 呼び出したトリガーを削除し翌日のトリガーを生成
    delTrigger();
    setTrigger();

    let trgtSpreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let trgtSh = trgtSpreadSheet.getSheetByName(SHEET1_NAME);

    let dataLastRow = trgtSh.getLastRow();
    let trgtRng = trgtSh.getRange(1, 1, dataLastRow, 3);
    let trgtAry: any[] = trgtRng.getValues();

    // 配列のインデックスは「0」から始まるため「-1」
    const phraseColIndex: number = 1 - 1;
    const descriptionColIndex: number = 2 - 1;
    const repeatFlgColIndex: number = 3 - 1;
    
    let flgRng = trgtSh.getRange(1,3, dataLastRow,3);
    let flgAry: number[][] = flgRng.getValues();
    let flgCnt: number = 0;
    flgAry.forEach(function(value){
        if (value[0] === 1) {
            flgCnt += value[0];
        }
    });
    
    // フラグのない行を取得
    let noFlgTrgtAry: any[] = new Array;
    const flgNum: number = 1;

    let trgtRowIndex: number;
    // 1周していたらフラグをリセット
    if (flgCnt === dataLastRow) {
        for (let i = 0; i <= dataLastRow - 1; i++) {
            trgtAry[i][repeatFlgColIndex] = 0;
        }
        noFlgTrgtAry = trgtAry
        trgtRowIndex = Math.floor(Math.random() * noFlgTrgtAry.length);

    }　else {
        trgtAry.forEach(function(el, index) {
            if (el[repeatFlgColIndex] != flgNum) {
                noFlgTrgtAry.push([el[phraseColIndex], el[descriptionColIndex], index]);
            }
        });
        trgtRowIndex = Math.floor(Math.random() * noFlgTrgtAry.length);
        const trgtIndexColIndex: number = 2;
        trgtRowIndex = noFlgTrgtAry[trgtRowIndex][trgtIndexColIndex];
    }
    
    // 返したフレーズ行にフラグをたてる
    trgtAry[trgtRowIndex][repeatFlgColIndex] = flgNum;
    // フラグをシートに反映
    trgtRng.setValues(trgtAry);
    
    let trgtPhrase: string = trgtAry[trgtRowIndex][phraseColIndex];
    let trgtDescription: string = trgtAry[trgtRowIndex][descriptionColIndex];

    let sendComment: string = `\`${ trgtPhrase }\`

${ trgtDescription }`
    
    PostMessageToSlack(sendComment);
    
    return ContentService.createTextOutput();
}

function setTrigger {
    let trgtTime: Date = new Date();
    // 実行時間翌日の指定時間を取得する
    trgtTime.setDate(trgtTime.getDate() + 1)
    trgtTime.setHours(Number(TARGETHOUR));
    trgtTime.setMinutes(Number(TARGETMINITUE));

    ScriptApp.newTrigger(TRIGGERTARGETFUNCTION).timeBased().at(trgtTime).create();
}
function delTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TRIGGERTARGETFUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}