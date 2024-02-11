function createCourses() {
  // スプレッドシート上で指定したクラス名を用いてクラスを一括作成する。

  // スプレッドシートの内容を2次元配列に格納する。
  const firstCell = SpreadsheetApp.getActiveSheet().getRange(1, 1);
  const values = firstCell.getDataRegion().getValues();
  if (firstCell.getValue() === "" || values.length === 1) {
    Browser.msgBox(
      "新しく作成するクラスの情報を入力してから実行してください。\\nA列の name のみ必須です。",
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

  const res = Browser.msgBox(
    "現在のシートの内容を消去します。\\n実行しますか？",
    Browser.Buttons.OK_CANCEL
  );
  if (res === "cancel") {
    return false;
  }

  const headers = [
    "name",
    "section",
    "descriptionHeading",
    "description",
    "room",
    "ownerId",
    "courseState",
  ];
  const referenceUrl =
    "https://developers.google.com/classroom/reference/rest/v1/courses#Course.FIELDS-table";
  const sh = SpreadsheetApp.getActiveSheet();
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, headers.length + 2).setValue(referenceUrl);
}

function listCourses() {
  // 自分が所属しているクラスの一覧をスプレッドシート上に出力する。

  // 自分が所属しているクラスのプロパティのうち、スプレッドシート上に出力するものの名前。
  const properties = [
    "name",
    "id",
    "ownerId",
    "enrollmentCode",
    "creationTime",
    "updateTime",
    "courseState",
  ];

  // クラスを取得し、変数 properties で指定したプロパティのみを抽出する。
  const myCourses = Classroom.Courses.list();
  const courseProperties = myCourses.courses.map((course) =>
    properties.map((property) => course[property])
  );

  const sh = SpreadsheetApp.getActiveSheet();
  const currentDataRegion = sh.getRange(1, 1).getDataRegion();
  currentDataRegion.clear();
  sh.getRange(2, 1, sh.getLastRow()).removeCheckboxes();

  sh.getRange(1, 1).setValue("target");
  sh.getRange(1, 2, 1, properties.length).setValues([properties]);

  sh.getRange(
    2,
    2,
    courseProperties.length,
    courseProperties[0].length
  ).setValues(courseProperties);

  // 出力した情報の左端にチェックボックスを挿入する。
  // これは、クラスの一括アーカイブや一括削除に使用するもの。
  sh.getRange(2, 1, courseProperties.length, 1).insertCheckboxes();

  // 公式レファレンスの URL をセルに書いておく。みんな見てね。
  const referenceUrl =
    "https://developers.google.com/classroom/reference/rest/v1/courses#Course.FIELDS-table";
  sh.getRange(1, properties.length + 3).setValue(referenceUrl);
}

function archiveCourses() {
  // 以下の関数に `ARCHIVE` という文字列を渡すことで、クラスの一括アーカイブを実行する。
  invokeArchiveOrRemovecourses("ARCHIVE");
}

function removeCourses() {
  const res = Browser.msgBox(
    "クラスを一括削除します。\\n実行しますか？",
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

  // セルの内容を2次元配列に格納する。
  // 2次元配列の1次元目の要素数が 1、つまり見出し行しかなかった場合、実行をキャンセルする。
  const firstCell = SpreadsheetApp.getActiveSheet().getRange(1, 1);
  const values = firstCell.getDataRegion().getValues();
  if (firstCell.getValue() === "" || values.length === 1) {
    Browser.msgBox(
      "まず初めにクラスの一覧を作成し、それから実行してください。",
      Browser.Buttons.OK
    );
    return false;
  }

  // 2次元配列をオブジェクトの配列に変換する。
  const requests = newObjectsFrom2DArray(values);

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
      "以下のクラスを" +
        actionName +
        "できませんでした。\\nスプレッドシート上のクラス名や ID が実際のものと一致しませんでした。\\nセルの内容が書き換わっている可能性があります。\\nクラスリストを再度作成してから実行してください。" +
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
      "以下のクラスを" +
        actionName +
        'できませんでした。\\nクラスの状態が "PROVISIONED" です。\\nクラスへの招待を承諾するか辞退してから再度実行してください。' +
        "\\n\\n" +
        provisionCourseNames
    );
  }
}

function createInvitations() {
  // セルの内容を2次元配列に格納する。
  // 2次元配列の1次元目の要素数が 1、つまり見出し行しかなかった場合、実行をキャンセルする。
  const firstCell = SpreadsheetApp.getActiveSheet().getRange(1, 1);
  const values = firstCell.getDataRegion().getValues();
  if (firstCell.getValue() === "" || values.length === 1) {
    Browser.msgBox(
      "以下の通りに情報を入力してから実行してください。A列以外は全て必須です。\\nA列: クラス名\\nB列: クラス ID\\nC列: ユーザー ID\\nD列: 役割 (STUDENT/TEACHER)",
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

  const res = Browser.msgBox(
    "現在のシートの内容を消去します。\\n実行しますか？",
    Browser.Buttons.OK_CANCEL
  );
  if (res === "cancel") {
    return false;
  }

  const headers = ["courseName", "courseId", "userId", "role"];
  const referenceUrl =
    "https://developers.google.com/classroom/reference/rest/v1/invitations#Invitation.FIELDS-table";
  const sh = SpreadsheetApp.getActiveSheet();
  sh.getRange(1, 1).getDataRegion().clearContent();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, headers.length + 2).setValue(referenceUrl);
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
    .addItem("クラスを一覧表示", "listCourses")
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
}
