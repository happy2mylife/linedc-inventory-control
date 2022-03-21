// 指定してください-----------------
const SHEET_URL = Config.SpreadSheetUrl;
const CHANNEL_ACCESS_TOKEN = Config.LineChanleAccessToken;
//--------------------------------

const LINE_REPLY_URL = "https://api.line.me/v2/bot/message/reply";
const LINE_API_PROFILE = "https://api.line.me/v2/bot/profile/";

// SpreadsheetのURL
const SpreadSheet = SpreadsheetApp.openByUrl(SHEET_URL);
const Sheet = {
  Item: SpreadSheet.getSheetByName("stock"),
  Order: SpreadSheet.getSheetByName("order"),
  User: SpreadSheet.getSheetByName("user"),
  Log: SpreadSheet.getSheetByName("log"),
  Status: SpreadSheet.getSheetByName("status"),
};

const StatusColumn = { Operation: 1, Item: 2, Status: 3 };
const Operation = {
  List: "在庫一覧",
  Add: "追加",
  Sub: "引出",
  Order: "発注",
};

const ItemColumnArrayIndex = { Name: 0, StockQuantity: 1 };
const StatusRowIndex = 1;

/**
 * LINEからリクエストを処理
 */
function doPost(request) {
  const receiveJSON = JSON.parse(request.postData.contents);
  // 詳細なデータ構造は以下を参照
  // https://developers.line.biz/ja/reference/messaging-api/#webhook-event-objects
  const event = receiveJSON.events[0];

  if (event.type == "follow") {
    setOperation(event.type);
    clearOperationItem();

    // アクセスユーザーを登録
    addUser(event.source.userId);

    replyToUser(
      event.replyToken,
      "ワイは在庫管理システムです。操作を選択してください。"
    );
    return;
  }

  // テキスト以外が送られてきたときは何もしない。
  if (event.message.type != "text") {
    return;
  }

  // 操作が指定された場合の応答を返す
  if (
    event.message.text == Operation.List ||
    event.message.text == Operation.Add ||
    event.message.text == Operation.Sub ||
    event.message.text == Operation.Order
  ) {
    doPrepareOperation(event);
    return;
  }

  const operation = getCurrentOperation();
  if (operation == Operation.Add || operation == Operation.Sub) {
    // 在庫追加or引出
    doAddOrSub(event, operation);
    return;
  } else if (operation == Operation.Order) {
    // 発注
    doOrder(event);
    return;
  }

  // 不明な操作時
  doUnknownOperation(event);
}

function setOperation(operation) {
  Sheet.Status.getRange(StatusRowIndex, StatusColumn.Operation).setValue(
    operation
  );
}

function clearOperation() {
  Sheet.Status.getRange(StatusRowIndex, StatusColumn.Operation).clearContent();
}

function clearOperationItem() {
  Sheet.Status.getRange(StatusRowIndex, StatusColumn.Item).clearContent();
}

function setItem(name) {
  Sheet.Status.getRange(StatusRowIndex, StatusColumn.Item).setValue(name);
}

/**
 * 現在処理中の操作名を取得
 * @returns シート記録の操作名
 */
function getCurrentOperation() {
  return Sheet.Status.getRange(
    StatusRowIndex,
    StatusColumn.Operation
  ).getValue();
}

function getCurrentItem() {
  return Sheet.Status.getRange(StatusRowIndex, StatusColumn.Item).getValue();
}

/**
 * 操作が指定された場合の処理
 *
 * @param event
 * @returns
 */
function doPrepareOperation(event) {
  let message;

  setOperation(event.message.text);
  clearOperationItem();

  if (event.message.text == Operation.List) {
    message = getItemList();
  } else if (event.message.text == Operation.Add) {
    message = "追加する商品名を送信してください。";
  } else if (event.message.text == Operation.Sub) {
    message = "引出す商品名を送信してください。";
  } else if (event.message.text == Operation.Order) {
    message = "発注する商品名を送信してください。";
  }

  replyToUser(event.replyToken, message);
}

/**
 * 不明な操作時の処理
 *
 * @param event
 */
function doUnknownOperation(event) {
  clearOperation();
  clearOperationItem();
  replyToUser(event.replyToken, "操作できませんでした。");
}

/**
 * 在庫追加or引出時の処理
 *
 * @param event
 * @returns
 */
function doAddOrSub(event, operation) {
  const operationItem = getCurrentItem();

  // 操作対象商品送信時
  if (!operationItem) {
    // 追加対象商品名が指定された場合の処理
    const rowId = findItemRowId(event.message.text);
    if (rowId == -1) {
      return replyToUser(
        event.replyToken,
        "管理されていない商品です。商品名を送信してください。"
      );
    }

    setItem(event.message.text);

    const message = operation == Operation.Add ? "追加する" : "引出す";
    replyToUser(event.replyToken, message + "数を送信してください。");
  } else {
    // 在庫増減数が指定された場合の処理

    let num = parseInt(event.message.text);
    if (isNaN(num)) {
      return replyToUser(
        event.replyToken,
        "入力値が不正です。数値を送信してください。"
      );
    }

    if (operation == Operation.Sub) {
      //　引出しの場合はマイナスにする
      num = -num;
    }

    // 操作対象の行を取得
    const rowId = findItemRowId(operationItem);
    const updatedNum = updateItemStock(num, rowId);
    if (updatedNum < 0) {
      return replyToUser(
        event.replyToken,
        "在庫数が足りないため引出せません。"
      );
    }

    // 追加・引出の操作ログを記録
    addOperationLog(event.source.userId, operationItem, operation, num);

    clearOperation();
    clearOperationItem();

    replyToUser(
      event.replyToken,
      "ワイは在庫管理システムです。" +
        "在庫数：" +
        operationItem +
        "(" +
        updatedNum +
        ")"
    );
  }
}

/**
 * 発注時の処理
 *
 * @param event
 * @returns
 */
function doOrder(event) {
  const operationItem = getCurrentItem();

  // 操作対象商品送信時
  if (!operationItem) {
    // 追加対象商品名が指定された場合の処理
    const rowId = findItemRowId(event.message.text);
    if (rowId == -1) {
      return replyToUser(
        event.replyToken,
        "管理されていない商品です。商品名を送信してください。"
      );
    }

    setItem(event.message.text);
    replyToUser(event.replyToken, "発注する数を送信してください。");
  } else {
    // 発注数が指定された場合の処理

    let orderNum = parseInt(event.message.text);
    if (isNaN(orderNum) || orderNum <= 0) {
      return replyToUser(
        event.replyToken,
        "入力値が不正です。数値を送信してください。"
      );
    }

    orderItem(orderNum, operationItem, event.source.userId);

    clearOperation();
    clearOperationItem();

    replyToUser(
      event.replyToken,
      "ワイは在庫管理システムです。" +
        operationItem +
        "(" +
        orderNum +
        ")を発注しました。"
    );
  }
}

/**
 * 在庫一覧
 */
function getItemList() {
  const items = Sheet.Item.getDataRange().getValues();
  let message = "";

  // ヘッダー行は処理しない
  for (let i = 1; i < items.length; i++) {
    // TODO: 1メッセージの上限注意
    message +=
      items[i][ItemColumnArrayIndex.Name] +
      "(" +
      items[i][ItemColumnArrayIndex.StockQuantity] +
      ")";
    if (i < items.length - 1) {
      message += "\n";
    }
  }

  return message;
}

/**
 * 追加・引出の履歴を記録
 */
function addOperationLog(userId, operationItem, operation, num) {
  const user = findUser(userId);
  let name = "不明";
  if (user != null) {
    name = user.displayName;
  }

  Sheet.Log.appendRow([getCurrentTime(), name, operationItem, operation, num]);
}

/**
 * 指定行の在庫数を更新
 *
 * @return 更新後の在庫数, -1は更新NG
 */
function updateItemStock(num, itemRowId) {
  // 対象商品の在庫数を取得
  const currentNum = Sheet.Item.getRange(itemRowId, 2).getValue();
  const updateNum = parseInt(currentNum) + num;
  if (updateNum < 0) {
    return -1;
  }
  Sheet.Item.getRange(itemRowId, 2).setValue(updateNum);

  return updateNum;
}

/**
 * 操作対象商品を検索
 *
 * @return 操作対象行番号
 */
function findItemRowId(name) {
  const itemsListRange = Sheet.Item.getRange(
    1,
    1,
    Sheet.Item.getLastRow()
  ).getValues();

  // 2次元配列を1次元配列に変換
  const itemsList = Array.prototype.concat.apply([], itemsListRange);
  const index = itemsList.lastIndexOf(name);

  // +1 = 行番号を同じにしている(-1であれば該当なし)
  return index === -1 ? index : index + 1;
}

/**
 * 発注数
 *
 * @param orderNum
 */
function orderItem(orderNum, orderItem, userId) {
  // usersシートからユーザー名を取得
  const user = findUser(userId);
  if (user == null) {
    return;
  }

  // orderシートに発注操作を記録
  Sheet.Order.appendRow([
    getCurrentTime(),
    user.displayName,
    orderItem,
    orderNum,
  ]);
}

/**
 * ユーザー追加
 * @param userId
 * @returns ユーザー名
 */
function addUser(userId) {
  const user = findUser(userId);
  if (user != null) {
    // 既に登録済み
    return;
  }

  // 新規アクセス
  const userName = getUserDisplayName(userId);
  Sheet.User.appendRow([userId, userName, getCurrentTime()]);

  return;
}

/**
 * 指定のユーザーIDのユーザー情報を取得
 * @param userId
 * @returns
 */
function findUser(userId) {
  const userValues = Sheet.User.getDataRange().getValues();
  const users = userValues.map((row) => {
    return { id: row[0], displayName: row[1], date: row[2] };
  });

  const index = users.findIndex((user) => {
    return user.id === userId;
  });

  return index != -1 ? users[index] : null;
}

/**
 * ユーザーのディスプレイ名を取得
 * @param userId
 * @returns
 */
function getUserDisplayName(userId) {
  var url = LINE_API_PROFILE + userId;
  var userProfile = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
  });
  return JSON.parse(userProfile).displayName;
}

/**
 * 現在時刻を取得
 * @returns
 */
function getCurrentTime() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
}

/**
 * クイックリプライメッセージを作成
 *
 * @return クイックリプライメッセージitem
 */
function createQuickReply() {
  return {
    items: [
      {
        type: "action",
        action: {
          type: "message",
          label: Operation.List,
          text: Operation.List,
        },
      },
      {
        type: "action",
        action: {
          type: "message",
          label: Operation.Add,
          text: Operation.Add,
        },
      },
      {
        type: "action",
        action: {
          type: "message",
          label: Operation.Sub,
          text: Operation.Sub,
        },
      },
      {
        type: "action",
        action: {
          type: "message",
          label: Operation.Order,
          text: Operation.Order,
        },
      },
      {
        type: "action",
        action: {
          type: "uri",
          label: "LIFF",
          uri: Config.LiffURL,
        },
      },
    ],
  };
}

/**
 * 該当ユーザーへ応答を返す
 */
function replyToUser(replyToken, text) {
  const replyText = {
    replyToken: replyToken,
    messages: [
      {
        type: "text",
        text: text,
        quickReply: createQuickReply(),
      },
    ],
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
    payload: JSON.stringify(replyText),
  };

  // Line該当ユーザーに応答を返している
  UrlFetchApp.fetch(LINE_REPLY_URL, options);
}
