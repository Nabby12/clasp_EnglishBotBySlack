const SLACK_VERIFICATIONTOKEN: string = PropertiesService.getScriptProperties().getProperty('SLACK_VERIFICATIONTOKEN');
const SLACK_WEBHOOK_URL: string = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
const SPREADSHEET_ID: string = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const SHEET1_NAME: string = PropertiesService.getScriptProperties().getProperty('SHEET1_NAME');
const SENDCOMMENTSTR: string = PropertiesService.getScriptProperties().getProperty('SENDCOMMENTSTR');

function doPost(e: string) {
    let verificationToken: string = e.parameter.token;
    if (verificationToken != SLACK_VERIFICATIONTOKEN) {
        throw new Error('Invalid Token');
    }
    
    let arg: string = e.parameter.text.replace('　', ' ').trim();

    let trgtSpreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let trgtSh = trgtSpreadSheet.getSheetByName(SHEET1_NAME);

    let dataLastRow = trgtSh.getLastRow();
    let trgtRng = trgtSh.getRange(1, 1, dataLastRow, 2);
    let trgtAry: string[] = trgtRng.getValues();
    let trgtRowIndex: number = Math.floor(Math.random() * dataLastRow);

    // 配列のインデックスは「0」から始まるため「-1」
    const phraseColIndex: number = 1 - 1;
    const descriptionColIndex: number = 2 - 1;
    
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

                if (el[phraseColIndex].indexOf(trgtWord) === -1) {
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
    }
    
    let trgtPhrase: string = trgtAry[trgtRowIndex][phraseColIndex];
    let trgtDescription: string = trgtAry[trgtRowIndex][descriptionColIndex];

    let sendComment: string;
    if (noMatchFlg === true) {
        sendComment = `\`${ trgtPhrase }\`

${ trgtDescription }

${ SENDCOMMENTSTR }`
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