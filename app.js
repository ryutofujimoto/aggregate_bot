// チャネルアクセストークン
const ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("ACCESS_TOKEN");
const HEADERS = {
    "Content-Type": "application/json; charset=UTF-8",
    Authorization: "Bearer " + ACCESS_TOKEN,
};

// Googleスプレッドシート
const SHEET_ID = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
const SHEET = SpreadsheetApp.openById(SHEET_ID);

// 集計対象時間「おはよう」
const GOOD_MORNING_START_HOUR = Number(PropertiesService.getScriptProperties().getProperty("GOOD_MORNING_START_HOUR"));
const GOOD_MORNING_START_MINUTE = Number(PropertiesService.getScriptProperties().getProperty("GOOD_MORNING_START_MINUTE"));
const GOOD_MORNING_END_HOUR = Number(PropertiesService.getScriptProperties().getProperty("GOOD_MORNING_END_HOUR"));
const GOOD_MORNING_END_MINUTE = Number(PropertiesService.getScriptProperties().getProperty("GOOD_MORNING_END_MINUTE"));

// 集計対象時間「手帳」
const NOTE_START_HOUR = Number(PropertiesService.getScriptProperties().getProperty("NOTE_START_HOUR"));
const NOTE_START_MINUTE = Number(PropertiesService.getScriptProperties().getProperty("NOTE_START_MINUTE"));
const NOTE_END_HOUR = Number(PropertiesService.getScriptProperties().getProperty("NOTE_END_HOUR"));
const NOTE_END_MINUTE = Number(PropertiesService.getScriptProperties().getProperty("NOTE_END_MINUTE"));

// ポイント
const GOOD_MORNING_POINT = Number(PropertiesService.getScriptProperties().getProperty("GOOD_MORNING_POINT"));
const NOTE_POINT = Number(PropertiesService.getScriptProperties().getProperty("NOTE_POINT"));

// 起動コマンド
const COMMAND_GOD_MORNING = "おはよう";
const COMMAND_TODAY = "今日の集計結果";
const COMMAND_WEEK = "今週の集計結果";
const COMMAND_MONTH = "今月の集計結果";

// 絵文字
const COMMAND_MONTH_TYPE_LIST = "🍙";
const COMMAND_BOOK_EMOJI = ["📖", "📕", "📗", "📘", "📙", "📚", "🗒️", "📝"];
const NUMVER_EMOJIS = ["0️⃣", "1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣"];

//================================//
//================================//

// メッセージを返信する
function replyMessage(replyToken, message) {
    let url = "https://api.line.me/v2/bot/message/reply";
    let postData = {
        replyToken: replyToken,
        messages: [
            {
                type: "text",
                text: message,
            },
        ],
    };
    let options = {
        method: "POST",
        headers: HEADERS,
        payload: JSON.stringify(postData),
    };

    return UrlFetchApp.fetch(url, options);
}

// チャットIDを返す
function getChatId(webhookData) {
    let userId = webhookData.source.userId;
    let groupId = webhookData.source.groupId;
    let roomId = webhookData.source.roomId;

    if (typeof roomId != "undefined") {
        return roomId;
    } else if (typeof groupId != "undefined") {
        return groupId;
    } else {
        return userId;
    }
}

// 登録ユーザー名取得
function getUserName(userId) {
    let url = "https://api.line.me/v2/bot/profile/" + userId;

    let options = {
        method: "get",
        headers: HEADERS,
    };

    let response = UrlFetchApp.fetch(url, options);
    let profile = JSON.parse(response.getContentText());

    if (profile && profile.displayName) {
        return profile.displayName;
    } else {
        return "Unknown User";
    }
}

// ポイントを種別ごとに取得
function getPoint(message) {
    let point = 0;

    if (message.includes(COMMAND_GOD_MORNING)) {
        point = GOOD_MORNING_POINT;
    } else if (COMMAND_BOOK_EMOJI.some((item) => message.includes(item))) {
        point = NOTE_POINT;
    }

    return point;
}

// フォーマット日付
function formatDate(date, format) {
    format = format.replace(/YYYY/, date.getFullYear());
    format = format.replace(/MM/, date.getMonth() + 1);
    format = format.replace(/DD/, date.getDate());
    format = format.replace(/hh/, date.getHours());
    format = format.replace(/mm/, date.getMinutes());
    format = format.replace(/ss/, date.getSeconds());
    format = format.replace(/week/, (dayOfWeekStr = ["日", "月", "火", "水", "木", "金", "土"][date.getDay()]));

    return format;
}

// 開始時間判定「おはよう」
function checkAggregateTimeGoodMorning(timestamp) {
    if (
        (timestamp.getHours() === GOOD_MORNING_START_HOUR && timestamp.getMinutes() >= GOOD_MORNING_START_MINUTE) ||
        (timestamp.getHours() > GOOD_MORNING_START_HOUR && timestamp.getHours() < GOOD_MORNING_END_HOUR) ||
        (timestamp.getHours() === GOOD_MORNING_END_HOUR && timestamp.getMinutes() <= GOOD_MORNING_END_MINUTE)
    ) {
        return true;
    } else {
        return false;
    }
}

// 開始時間判定「手帳」
function checkAggregateTimeNote(timestamp) {
    if (
        (timestamp.getHours() === NOTE_START_HOUR && timestamp.getMinutes() >= NOTE_START_MINUTE) ||
        (timestamp.getHours() > NOTE_START_HOUR && timestamp.getHours() < NOTE_END_HOUR) ||
        (timestamp.getHours() === NOTE_END_HOUR && timestamp.getMinutes() <= NOTE_END_MINUTE)
    ) {
        return true;
    } else {
        return false;
    }
}

// ファイル出力メッセージ
function outputFileMessage(timestamp, senderId, message) {
    return SHEET.appendRow([timestamp, formatDate(timestamp, "YYYY年MM月DD日hh時mm分ss秒（week）"), message, senderId, getUserName(senderId)]);
}

// 今日の集計結果
function todayAggregateResult(replyToken) {
    let data = SHEET.getDataRange().getValues();
    let currentDate = new Date();
    let summary = {};

    for (let i = 0; i < data.length; i++) {
        let timestamp = new Date(data[i][0]);
        let getMessage = data[i][2];
        let row = data[i];
        let userName = row[4];

        if (
            timestamp.getDate() !== currentDate.getDate() ||
            timestamp.getMonth() !== currentDate.getMonth() ||
            timestamp.getYear() !== currentDate.getYear()
        ) {
            // 年月日が一致しない場合はコンテニュー
            continue;
        }

        if (summary[userName]) {
            summary[userName] += getPoint(getMessage);
        } else {
            summary[userName] = getPoint(getMessage);
        }
    }

    let reply = "今日の「おはよう」「" + COMMAND_BOOK_EMOJI + "」メッセージの一覧\n";

    summary = Object.entries(summary).sort((a, b) => b[1] - a[1]);

    summary.forEach(function (entry, index) {
        const countValue = convertionDisplayCount(entry[1]);
        reply += index + 1 + ". " + entry[0] + ": " + countValue + "ポイント\n";
    });

    return replyMessage(replyToken, reply);
}

// 今週の集計結果
function weekAggregateResult(replyToken) {
    let data = SHEET.getDataRange().getValues();
    let currentDate = new Date();
    let summary = {};

    // 今週の日曜日を計算
    let sunday = new Date();
    sunday.setHours(0, 0, 0, 0);
    sunday.setDate(currentDate.getDate() - currentDate.getDay());

    for (let i = 0; i < data.length; i++) {
        let timestamp = new Date(data[i][0]);
        let getMessage = data[i][2];
        let row = data[i];
        let userName = row[4];

        // 今週の日曜日以降
        if (timestamp >= sunday) {
            if (summary[userName]) {
                summary[userName] += getPoint(getMessage);
            } else {
                summary[userName] = getPoint(getMessage);
            }
        }
    }

    let reply = "今週の「おはよう」「" + COMMAND_BOOK_EMOJI + "」メッセージの一覧\n";

    summary = Object.entries(summary).sort((a, b) => b[1] - a[1]);

    summary.forEach(function (entry, index) {
        const countValue = convertionDisplayCount(entry[1]);
        reply += index + 1 + ". " + entry[0] + ": " + countValue + "ポイント\n";
    });

    return replyMessage(replyToken, reply);
}

// 今月の集計結果
function monthAggregateResult(replyToken) {
    let data = SHEET.getDataRange().getValues();
    let currentDate = new Date();
    let summary = {};

    for (let i = 0; i < data.length; i++) {
        let timestamp = new Date(data[i][0]);
        let getMessage = data[i][2];
        let row = data[i];
        let userName = row[4];

        if (timestamp.getMonth() === currentDate.getMonth()) {
            if (summary[userName]) {
                summary[userName] += getPoint(getMessage);
            } else {
                summary[userName] = getPoint(getMessage);
            }
        }
    }

    let reply = "今月の「おはよう」「" + COMMAND_BOOK_EMOJI + "」メッセージの一覧\n";

    summary = Object.entries(summary).sort((a, b) => b[1] - a[1]);

    summary.forEach(function (entry, index) {
        const countValue = convertionDisplayCount(entry[1]);
        reply += index + 1 + ". " + entry[0] + ": " + countValue + "ポイント\n";
    });

    return replyMessage(replyToken, reply);
}

// ユーザー別今月の集計結果
function monthUserListAggregateResult(replyToken) {
    let data = SHEET.getDataRange().getValues();
    let currentDate = new Date();
    let goodMorningSummary = {};
    let bookSummary = {};

    for (let i = 0; i < data.length; i++) {
        let timestamp = new Date(data[i][0]);
        let getMessage = data[i][2];
        let row = data[i];
        let userName = row[4];

        if (timestamp.getMonth() !== currentDate.getMonth()) {
            continue;
        }

        let weekStart = new Date(timestamp);
        weekStart.setHours(0, 0, 0, 0);
        weekStart.setDate(weekStart.getDate() - weekStart.getDay());

        if (weekStart <= timestamp) {
            let weekStartStr = Utilities.formatDate(weekStart, "GMT+09:00", "MM/dd");
            let weekEnd = new Date(weekStart);
            weekEnd.setDate(weekStart.getDate() + 6);
            let weekEndStr = Utilities.formatDate(weekEnd, "GMT+09:00", "MM/dd");

            let weekLabel = weekStartStr + " ～ " + weekEndStr;

            // 「おはよう」集計
            if (getMessage.includes(COMMAND_GOD_MORNING)) {
                if (!goodMorningSummary[userName]) {
                    goodMorningSummary[userName] = {};
                }

                if (!goodMorningSummary[userName][weekLabel]) {
                    goodMorningSummary[userName][weekLabel] = getPoint(getMessage);
                } else {
                    goodMorningSummary[userName][weekLabel] += getPoint(getMessage);
                }
            }

            // 手帳集計
            if (COMMAND_BOOK_EMOJI.some((item) => getMessage.includes(item))) {
                if (!bookSummary[userName]) {
                    bookSummary[userName] = {};
                }

                if (!bookSummary[userName][weekLabel]) {
                    bookSummary[userName][weekLabel] = getPoint(getMessage);
                } else {
                    bookSummary[userName][weekLabel] += getPoint(getMessage);
                }
            }
        }
    }

    let reply = "今月の「おはよう」メッセージの一覧\n";
    for (let user in goodMorningSummary) {
        reply += user + ":\n";
        for (let weekLabel in goodMorningSummary[user]) {
            const countValue = convertionDisplayCount(goodMorningSummary[user][weekLabel]);
            reply += weekLabel + ": " + countValue + "ポイント\n";
        }
    }

    reply += "-------------------------------\n";
    reply += "今月の「手帳」メッセージの一覧\n";
    for (let user in bookSummary) {
        reply += user + ":\n";
        for (let weekLabel in bookSummary[user]) {
            const countValue = convertionDisplayCount(bookSummary[user][weekLabel]);
            reply += weekLabel + ": " + countValue + "ポイント\n";
        }
    }

    return replyMessage(replyToken, reply);
}

// 集計対象時間を表示用に変換
function convertionDisplayTime(startHour, startMinute, endHour, endMinute) {
    return "AM" + startHour + ":" + startMinute + " ~ AM" + endHour + ":" + endMinute;
}

// 回数表示を絵文字の数字に変換
function convertionDisplayCount(count) {
    const constString = count
        .toString()
        .split("")
        .map((number) => NUMVER_EMOJIS[number])
        .join("");
    return constString;
}

//================================//
//================================//

// main処理
function doPost(e) {
    let webhookData = JSON.parse(e.postData.contents).events[0];
    let replyToken = webhookData.replyToken;
    let message = webhookData.message.text;
    let currentTime = new Date(webhookData.timestamp);
    let SenderID = webhookData.source.userId;

    let chatId = getChatId(webhookData);

    let displayTimeGoodMorning = convertionDisplayTime(
        GOOD_MORNING_START_HOUR,
        GOOD_MORNING_START_MINUTE,
        GOOD_MORNING_END_HOUR,
        GOOD_MORNING_END_MINUTE
    );
    let displayTimeNote = convertionDisplayTime(NOTE_START_HOUR, NOTE_START_MINUTE, NOTE_END_HOUR, NOTE_END_MINUTE);

    // ヘルプコマンド
    let help1 = "ヘルプ";
    let help2 = "help";
    if (message.includes(help1) || message.includes(help2)) {
        let reply = "集計時間：" + displayTimeGoodMorning;
        reply += "\n「" + COMMAND_GOD_MORNING + "」メッセージを集計します。";
        reply += "\n-------------------------------";
        reply += "\n集計時間：" + displayTimeNote;
        reply += "\n「" + COMMAND_BOOK_EMOJI + "」";
        reply += "\nメッセージを集計します。";
        reply += "\n-------------------------------";
        reply += "\n👇👇集計結果表示一覧👇👇";
        reply += "\n今日：「" + COMMAND_TODAY + "」";
        reply += "\n今週：「" + COMMAND_WEEK + "」";
        reply += "\n今月：「" + COMMAND_MONTH + "」";
        reply += "\n隠し：「" + COMMAND_MONTH_TYPE_LIST + "」";

        return replyMessage(replyToken, reply);
    }

    // 今日の集計結果
    if (message.includes(COMMAND_TODAY)) {
        todayAggregateResult(replyToken);
    }

    // 今週の集計結果
    if (message.includes(COMMAND_WEEK)) {
        weekAggregateResult(replyToken);
    }

    // 今月の集計結果
    if (message.includes(COMMAND_MONTH)) {
        monthAggregateResult(replyToken);
    }

    // 今月の種類別一覧集計結果
    if (message.includes(COMMAND_MONTH_TYPE_LIST)) {
        monthUserListAggregateResult(replyToken);
    }

    // おはようメッセージをスプレッドシートに保存
    if (message.includes(COMMAND_GOD_MORNING)) {
        if (!checkAggregateTimeGoodMorning(currentTime)) {
            return replyMessage(
                replyToken,
                "「" + COMMAND_GOD_MORNING + "」メッセージ集計時間外です。\n集計時間は" + displayTimeGoodMorning + "です"
            );
        }

        return outputFileMessage(currentTime, SenderID, message);
    }

    // 手帳メッセージをスプレッドシートに保存
    if (COMMAND_BOOK_EMOJI.some((item) => message.includes(item))) {
        if (!checkAggregateTimeNote(currentTime)) {
            return replyMessage(replyToken, "「" + COMMAND_BOOK_EMOJI + "」メッセージ集計時間外です。\n集計時間は" + displayTimeNote + "です");
        }
        return outputFileMessage(currentTime, SenderID, message);
    }
}
