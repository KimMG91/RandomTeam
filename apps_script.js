// ============================================
// Google Apps Script - 구글 시트에 붙여넣기
// 시트 > 확장 프로그램 > Apps Script > 이 코드 붙여넣기 > 배포 > 웹앱으로 배포
// 배포 시 "누구나 액세스 가능"으로 설정
// ============================================

function doGet(e) {
    var action = e.parameter.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var callback = e.parameter.callback;

    if (action === 'read') {
        var result = readData(ss);
        if (callback) {
            return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
                .setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return sendJson(result);
    }

    if (action === 'save') {
        var data = JSON.parse(e.parameter.data);
        var result = saveData(ss, data);
        if (callback) {
            return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
                .setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return sendJson(result);
    }

    return sendJson({ error: 'invalid action' });
}

function doPost(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data;

    // form submit 방식 (parameter.data) 또는 fetch body 방식 지원
    if (e.parameter && e.parameter.data) {
        data = JSON.parse(e.parameter.data);
    } else {
        data = JSON.parse(e.postData.contents);
    }

    if (data.action === 'save') {
        return sendJson(saveData(ss, data));
    }

    return sendJson({ error: 'invalid action' });
}

function sendJson(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}

// 데이터 읽기
function readData(ss) {
    var result = { members: [], teams: [] };

    var memberSheet = ss.getSheetByName('멤버');
    if (memberSheet && memberSheet.getLastRow() > 1) {
        var memberData = memberSheet.getDataRange().getValues();
        for (var i = 1; i < memberData.length; i++) {
            var row = memberData[i];
            if (!row[0]) continue;
            result.members.push({
                name: String(row[0]).trim(),
                ageGroup: String(row[1]).trim(),
                gender: row[2] === '여' ? 'female' : 'male'
            });
        }
    }

    var teamSheet = ss.getSheetByName('팀설정');
    if (teamSheet && teamSheet.getLastRow() > 1) {
        var teamData = teamSheet.getDataRange().getValues();
        for (var i = 1; i < teamData.length; i++) {
            var row = teamData[i];
            if (!row[0]) continue;
            result.teams.push({
                name: String(row[0]).trim(),
                color: row[1] ? String(row[1]).trim() : '#cccccc',
                limit: row[2] ? parseInt(row[2]) : ''
            });
        }
    }

    // 고정조 시트 읽기
    var fixedSheet = ss.getSheetByName('고정조');
    result.fixedGroups = [];
    if (fixedSheet && fixedSheet.getLastRow() > 1) {
        var fixedData = fixedSheet.getDataRange().getValues();
        for (var i = 1; i < fixedData.length; i++) {
            var row = fixedData[i];
            var group = [];
            for (var j = 0; j < row.length; j++) {
                if (row[j]) group.push(String(row[j]).trim());
            }
            if (group.length >= 2) result.fixedGroups.push(group);
        }
    }

    return result;
}

// 데이터 저장
function saveData(ss, data) {
    if (data.members) {
        var memberSheet = ss.getSheetByName('멤버');
        if (!memberSheet) memberSheet = ss.insertSheet('멤버');
        memberSheet.clear();
        var memberRows = [['이름', '연령대', '성별']];
        data.members.forEach(function(m) {
            memberRows.push([m.name, m.ageGroup, m.gender === 'female' ? '여' : '남']);
        });
        memberSheet.getRange(1, 1, memberRows.length, 3).setValues(memberRows);
    }

    if (data.teams) {
        var teamSheet = ss.getSheetByName('팀설정');
        if (!teamSheet) teamSheet = ss.insertSheet('팀설정');
        teamSheet.clear();
        var teamRows = [['팀이름', '색상', '인원제한']];
        data.teams.forEach(function(t) {
            teamRows.push([t.name, t.color, t.limit || '']);
        });
        teamSheet.getRange(1, 1, teamRows.length, 3).setValues(teamRows);
    }

    return { success: true };
}
