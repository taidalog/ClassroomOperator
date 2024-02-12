const helpURL =
  "https://github.com/taidalog/ClassroomOperator/blob/v1.0.0/README.ja.md";

const creationSheetInfo = {
  name: "作成",
  headers: [
    "name",
    "section",
    "descriptionHeading",
    "description",
    "room",
    "ownerId",
    "courseState",
  ],
  referenceUrl:
    "https://developers.google.com/classroom/reference/rest/v1/courses#Course.FIELDS-table",
};

const listSheetInfo = {
  name: "一覧・削除",
  headers: [
    "name",
    "id",
    "ownerId",
    "enrollmentCode",
    "creationTime",
    "updateTime",
    "courseState",
  ],
  referenceUrl:
    "https://developers.google.com/classroom/reference/rest/v1/courses#Course.FIELDS-table",
};

const listSheetInfoForNew = {
  name: listSheetInfo.name,
  headers: ["target"].concat(listSheetInfo.headers),
  referenceUrl: listSheetInfo.referenceUrl,
};

const invitationSheetInfo = {
  name: "招待",
  headers: ["courseName", "courseId", "userId", "role"],
  referenceUrl:
    "https://developers.google.com/classroom/reference/rest/v1/invitations#Invitation.FIELDS-table",
};

function newSheet(sheetInfo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    sheetInfo.name
  );

  if (null !== sh) {
    return false;
  }

  const newsh = SpreadsheetApp.getActiveSpreadsheet()
    .insertSheet()
    .setName(sheetInfo.name);
  newsh
    .getRange(1, 1, 1, sheetInfo.headers.length)
    .setValues([sheetInfo.headers]);
  newsh
    .getRange(1, sheetInfo.headers.length + 2)
    .setValue(sheetInfo.referenceUrl)
    .offset(1, 0)
    .setValue(helpURL);
  newsh.setFrozenRows(1);
}

const newCourseCreationSheet = () => newSheet(creationSheetInfo);
const newCourseListSheet = () => newSheet(listSheetInfoForNew);
const newInvitationSheet = () => newSheet(invitationSheetInfo);

function createCourses() {
  // スプレッドシート上で指定したクラス名を用いてクラスを一括作成する。

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(creationSheetInfo.name);

  if (null == sh) {
    newCourseCreationSheet();
    Browser.msgBox(
      "途中で終了しました",
      "[" +
        creationSheetInfo.name +
        "] シートが存在しなかったので作成しました。\\n新しく作成するクラスの情報を入力してから再度実行してください。\\nA列の name のみ必須です。",
      Browser.Buttons.OK
    );
    return false;
  }

  // スプレッドシートの内容を2次元配列に格納する。
  const firstCell = sh.getRange(1, 1);
  const values = firstCell.getDataRegion().getValues();

  if (firstCell.getValue() === "" || values.length === 1) {
    Browser.msgBox(
      "途中で終了しました",
      "[" +
        creationSheetInfo.name +
        "] シートに必要な情報が入力されていないため、処理を中断しました。\\n新しく作成するクラスの情報を入力してから再度実行してください。\\nA列の name のみ必須です。",
      Browser.Buttons.OK
    );
    return false;
  }

  // 2次元配列をオブジェクトの配列に変換する。
  const requests = newObjectsFrom2DArray(values);

  // `ownerId` または `courseState` が空欄だった場合、それぞれ "me" または "PROVISIONED" を指定してクラスを作成。
  requests.map((request) => {
    const resource = Object.assign({}, request);
    const ownerId = request.ownerId !== "" ? request.ownerId : "me";
    const courseState =
      request.courseState !== "" ? request.courseState : "PROVISIONED";
    resource.ownerId = ownerId;
    resource.courseState = courseState;
    Classroom.Courses.create(resource);
  });
}

function resetCourseCreationSheet() {
  // 選択中のシートの内容を消去し、クラス一括作成用の見出しを作成する。

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(creationSheetInfo.name);

  if (null == sh) {
    newCourseCreationSheet();
    Browser.msgBox(
      "終了しました",
      "[" +
        creationSheetInfo.name +
        "] シートが存在しなかったので作成しました。",
      Browser.Buttons.OK
    );
    return false;
  }

  const res = Browser.msgBox(
    "確認",
    "[" +
      creationSheetInfo.name +
      "] シートの内容を消去します。\\n実行しますか？",
    Browser.Buttons.OK_CANCEL
  );
  if (res === "cancel") {
    return false;
  }

  sh.clearContents();
  sh.getRange(1, 1, 1, creationSheetInfo.headers.length).setValues([
    creationSheetInfo.headers,
  ]);
  sh.getRange(1, creationSheetInfo.headers.length + 2)
    .setValue(creationSheetInfo.referenceUrl)
    .offset(1, 0)
    .setValue(helpURL);
}

function listCourses() {
  // 自分が所属しているクラスの一覧をスプレッドシート上に出力する。

  // クラスを取得し、listSheetInfo.headers で指定したプロパティのみを抽出する。
  const myCourses = Classroom.Courses.list();
  const courseProperties = myCourses.courses.map((course) =>
    listSheetInfo.headers.map((x) => course[x])
  );

  const getSheet = () => {
    const temp = ss.getSheetByName(listSheetInfo.name);
    if (null !== temp) {
      return temp;
    } else {
      newCourseListSheet();
      return ss.getSheetByName(listSheetInfo.name);
    }
  };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getSheet();

  const currentDataRegion = sh.getRange(1, 1).getDataRegion();
  currentDataRegion.clear();
  sh.getRange(2, 1, sh.getLastRow()).removeCheckboxes();

  sh.getRange(1, 1).setValue("target");
  sh.getRange(1, 2, 1, listSheetInfo.headers.length).setValues([
    listSheetInfo.headers,
  ]);

  sh.getRange(
    2,
    2,
    courseProperties.length,
    courseProperties[0].length
  ).setValues(courseProperties);

  // 出力した情報の左端にチェックボックスを挿入する。
  // これは、クラスの一括アーカイブや一括削除に使用するもの。
  sh.getRange(2, 1, courseProperties.length, 1).insertCheckboxes();

  // 公式レファレンスと README の URL をセルに書いておく。みんな見てね。
  sh.getRange(1, listSheetInfo.headers.length + 3)
    .setValue(listSheetInfo.referenceUrl)
    .offset(1, 0)
    .setValue(helpURL);
}

function archiveCourses() {
  const res = Browser.msgBox(
    "確認",
    "[" +
      listSheetInfo.name +
      "] シートでチェックが入っているクラスを一括アーカイブします。\\n実行しますか？",
    Browser.Buttons.OK_CANCEL
  );
  if (res === "cancel") {
    return false;
  }

  // 以下の関数に `ARCHIVE` という文字列を渡すことで、クラスの一括アーカイブを実行する。
  invokeArchiveOrRemovecourses("ARCHIVE");
}

function removeCourses() {
  const res = Browser.msgBox(
    "確認",
    "[" +
      listSheetInfo.name +
      "] シートでチェックが入っているクラスを一括削除します。\\n実行しますか？",
    Browser.Buttons.OK_CANCEL
  );
  if (res === "cancel") {
    return false;
  }

  // 以下の関数に `REMOVE` という文字列を渡すことで、クラスの一括削除を実行する。
  invokeArchiveOrRemovecourses("REMOVE");
}

function invokeArchiveOrRemovecourses(action) {
  // 引数として受け取った文字列に応じて、クラスの一括アーカイブ又は一括削除を実行する。

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(listSheetInfo.name);

  if (null == sh) {
    Browser.msgBox(
      "途中で終了しました",
      "[" +
        listSheetInfo.name +
        "] シートが存在しなかったので、処理を中断しました。\\nメニューから [クラスを一覧表示] を実行して、削除またはアーカイブするクラスにチェックを入れてから再度実行してください。",
      Browser.Buttons.OK
    );
    return false;
  }

  // セルの内容を2次元配列に格納する。
  // 2次元配列の1次元目の要素数が 1、つまり見出し行しかなかった場合、実行をキャンセルする。
  const firstCell = sh.getRange(1, 1);
  const values = firstCell.getDataRegion().getValues();
  if (firstCell.getValue() === "" || values.length === 1) {
    Browser.msgBox(
      "途中で終了しました",
      "[" +
        listSheetInfo.name +
        "] シートに必要な情報が入力されていないため、処理を中断しました。\\nメニューから [クラスを一覧表示] を実行して、削除またはアーカイブするクラスにチェックを入れてから再度実行してください。",
      Browser.Buttons.OK
    );
    return false;
  }

  // 2次元配列をオブジェクトの配列に変換する。
  const requests = newObjectsFrom2DArray(values);

  if (requests.filter((x) => x.target === true).length === 0) {
    Browser.msgBox(
      "途中で終了しました",
      "[" +
        listSheetInfo.name +
        "] シートのどのクラスにもチェックが入っていないので、処理を中断しました。\\n削除またはアーカイブするクラスにチェックを入れてから再度実行してください。",
      Browser.Buttons.OK
    );
    return false;
  }

  // チェックボックスにチェックが入ったものを抽出し、
  // その内容を関数で確認する。
  // 確認している内容については validateRequest 関数を参照のこと。
  const validationResults = requests
    .filter((request) => request.target === true)
    .map((request) => validateRequest(request));

  const actionName =
    action === "ARCHIVE" ? "アーカイブ" : action === "REMOVE" ? "削除" : "";

  // クラスを削除するには。いったんアーカイブする必要がある。
  // そのためのオブジェクト。
  const resource = { courseState: "ARCHIVED" };
  const updateMask = { updateMask: "courseState" };

  // アーカイブまたは削除を実行する。
  // 招待されているがまだ承諾も辞退もしていないクラス (PROVISIONED) はエラーになるため除外する。
  // 意図しないクラスをアーカイブ又は削除することを避けるため
  // スプレッドシート上のクラス名や ID が実際のものと一致しなかったクラスも除外する。
  validationResults
    .filter((validationResult) => validationResult.matched)
    .map((validationResult) => validationResult.course)
    .filter((course) => course.courseState !== "PROVISIONED")
    .map((course) => {
      Logger.log(actionName + "対象: " + newNameIdString(course) + ".");
      Classroom.Courses.patch(resource, course.id, updateMask);
      if (action === "REMOVE") {
        Classroom.Courses.remove(course.id);
      }
    });

  // スプレッドシート上のクラス名や ID が、実際のものと一致しなかったクラスを表示する。
  const unmatchedCourseNames = validationResults
    .filter((validationResult) => validationResult.matched === false)
    .map((validationResult) => newNameIdString(validationResult.request))
    .join("\\n");
  if (unmatchedCourseNames) {
    Browser.msgBox(
      "終了しました",
      "以下のクラスを" +
        actionName +
        "できませんでした。\\nスプレッドシート上のクラス名や ID が実際のものと一致しませんでした。\\nセルの内容が書き換わっている可能性があります。\\nクラスリストを作成しなおしてから再度実行してください。" +
        "\\n\\n" +
        unmatchedCourseNames
    );
  }

  // 招待されているがまだ承諾も辞退もしていないクラス (PROVISIONED) を表示する。
  const provisionCourseNames = validationResults
    .filter((validationResult) => validationResult.matched)
    .filter(
      (validationResult) =>
        validationResult.course.courseState === "PROVISIONED"
    )
    .map((validationResult) => newNameIdString(validationResult.course))
    .join("\\n");
  if (provisionCourseNames) {
    Browser.msgBox(
      "終了しました",
      "以下のクラスを" +
        actionName +
        'できませんでした。\\nクラスの状態が "PROVISIONED" です。\\nクラスへの招待を承諾するか辞退してから再度実行してください。' +
        "\\n\\n" +
        provisionCourseNames
    );
  }
}

function createInvitations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(invitationSheetInfo.name);

  if (null == sh) {
    newInvitationSheet();
    Browser.msgBox(
      "途中で終了しました",
      "[" +
        invitationSheetInfo.name +
        "] シートが存在しなかったので作成しました。\\n以下の通りに情報を入力してから再度実行してください。A列以外は全て必須です。\\nA列: クラス名\\nB列: クラス ID\\nC列: ユーザー ID\\nD列: 役割 (STUDENT/TEACHER)",
      Browser.Buttons.OK
    );
    return false;
  }

  // セルの内容を2次元配列に格納する。
  // 2次元配列の1次元目の要素数が 1、つまり見出し行しかなかった場合、実行をキャンセルする。
  const firstCell = sh.getRange(1, 1);
  const values = firstCell.getDataRegion().getValues();
  if (firstCell.getValue() === "" || values.length === 1) {
    Browser.msgBox(
      "途中で終了しました",
      "[" +
        invitationSheetInfo.name +
        "] シートに必要な情報が入力されていないため、処理を中断しました。\\n以下の通りに情報を入力してから再度実行してください。A列以外は全て必須です。\\nA列: クラス名\\nB列: クラス ID\\nC列: ユーザー ID\\nD列: 役割 (STUDENT/TEACHER)",
      Browser.Buttons.OK
    );
    return false;
  }

  // 2次元配列をオブジェクトの配列に変換する。
  const requests = newObjectsFrom2DArray(values);

  requests
    .map((request) => {
      return {
        courseId: String(request.courseId),
        userId: request.userId,
        role: request.role,
      };
    })
    .map((request) => Classroom.Invitations.create(request));
}

function resetInvitationSheet() {
  // 選択中のシートの内容を消去し、一括招待用の見出しを作成する。

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(invitationSheetInfo.name);

  if (null == sh) {
    newInvitationSheet();
    Browser.msgBox(
      "終了しました",
      "[" +
        invitationSheetInfo.name +
        "] シートが存在しなかったので作成しました。",
      Browser.Buttons.OK
    );
    return false;
  }

  const res = Browser.msgBox(
    "確認",
    "[" +
      invitationSheetInfo.name +
      "] シートの内容を消去します。\\n実行しますか？",
    Browser.Buttons.OK_CANCEL
  );
  if (res === "cancel") {
    return false;
  }

  sh.getRange(1, 1).getDataRegion().clearContent();
  sh.getRange(1, 1, 1, invitationSheetInfo.headers.length).setValues([
    invitationSheetInfo.headers,
  ]);
  sh.getRange(1, invitationSheetInfo.headers.length + 2)
    .setValue(invitationSheetInfo.referenceUrl)
    .offset(1, 0)
    .setValue(helpURL);
}

function tryCoursesGet(course_id) {
  try {
    const course = Classroom.Courses.get(course_id);
    return { succeed: true, course: course };
  } catch {
    return { succeed: false, course: undefined };
  }
}

function validateRequest(request) {
  const course = tryCoursesGet(request.id);
  return {
    request: request,
    course: course.course,
    matched:
      course.succeed &&
      course.course.name === request.name &&
      course.course.id == request.id,
  };
}

function newNameIdString(course) {
  return course.name + " (id: " + course.id + ")";
}

function newObjectsFrom2DArray(arr) {
  // 参考ページ
  // https://front-works.co.jp/blog/gas-creating-objects-from-spreadsheet/
  const [headers, ...records] = arr;
  return records.map((record) =>
    Object.fromEntries(record.map((value, i) => [headers[i], value]))
  );
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("クラスルーム一括処理")
    .addSubMenu(
      ui.createMenu("クラスを一覧表示").addItem("実行", "listCourses")
    )
    .addSubMenu(
      ui
        .createMenu("クラスを一括作成")
        .addItem("実行", "createCourses")
        .addItem("シート初期化", "resetCourseCreationSheet")
    )
    .addSubMenu(
      ui
        .createMenu("クラスに一括招待")
        .addItem("実行", "createInvitations")
        .addItem("シート初期化", "resetInvitationSheet")
    )
    .addSubMenu(
      ui
        .createMenu("クラスを一括削除")
        .addItem("一括アーカイブ", "archiveCourses")
        .addItem("一括削除", "removeCourses")
    )
    .addToUi();
  newCourseCreationSheet();
  newCourseListSheet();
  newInvitationSheet();
}
