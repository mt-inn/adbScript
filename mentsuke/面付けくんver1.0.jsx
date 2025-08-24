#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);
// === パス生成関数 ===
var mmToPt = function (mm) { return mm * 2.83465; };
function createTonbow (doc, obj) {
    if ( obj.typename == "GroupItem" && obj.clipped == true ) {
        var bounds = obj.pageItems[0].geometricBounds;
        var X = bounds[0];   // 左上のX座標
        var Y = bounds[1];   // 左上のY座標
        var W = bounds[2] - bounds[0];   // 横幅
        var H = bounds[1] - bounds[3];   // 縦幅
    } else {
        var position = obj.position;
        var X = position[0];   // 左上のX座標
        var Y = position[1];   // 左上のY座標
        var W = obj.width;   // 横幅
        var H = obj.height;   // 縦幅
    }

    var l = 34.02;
    var s = 25.11;
    var strokeW = 0.3; // 線幅（pt）
    var strokeRGB = new RGBColor();
    strokeRGB.red = 0;
    strokeRGB.green = 0;
    strokeRGB.blue = 0;
    var cY = Y - H / 2;
    var cX = X + W / 2;
    var tonbow = [];

    function createLine(x1, y1, x2, y2) {
        var path = doc.pathItems.add();
        path.setEntirePath([[x1, y1], [x2, y2]]);
        path.stroked = true;
        path.strokeWidth = strokeW;
        path.strokeColor = strokeRGB;
        path.filled = false;
        return path;
    }

    function createPolyline(points) {
        var path = doc.pathItems.add();
        path.setEntirePath(points);
        path.stroked = true;
        path.strokeWidth = strokeW;
        path.strokeColor = strokeRGB;
        path.filled = false;
        return path;
    }
    //トンボ作成
    //左上から時計回り、y座標が大きいのを先に
    //y は上が正、xは右が正
    // 左上
    tonbow.push(createPolyline([
        [X - l, Y - mmToPt(3)],
        [X, Y - mmToPt(3)],
        [X, Y - mmToPt(3)+ s]
    ]));//0
    tonbow.push(createPolyline([
        [X + mmToPt(3) - s, Y],
        [X + mmToPt(3), Y],
        [X + mmToPt(3), Y + l]
    ]));//1

    //上辺中央
    tonbow.push(createLine(cX, Y + 12.15 - mmToPt(3), cX, Y + 36.15 - mmToPt(3)));//2
    tonbow.push(createLine(cX - 36  , Y + 18.15 - mmToPt(3), cX + 36, Y + 18.15 - mmToPt(3)));//3

    //右上
    tonbow.push(createPolyline([
        [X + W, Y - mmToPt(3) + s],
        [X + W, Y - mmToPt(3)],
        [X + W + l, Y - mmToPt(3)]
    ]));//4
    tonbow.push(createPolyline([
        [X + W - mmToPt(3), Y + l],
        [X + W - mmToPt(3), Y],
        [X + W - mmToPt(3) + s, Y]
    ]));//5

    //右辺中央
    tonbow.push(createLine(X + W + 36.15 - mmToPt(3), cY, X + W + 12.15 - mmToPt(3), cY));//6
    tonbow.push(createLine(X + W + 18.15 - mmToPt(3), cY + 36, X + W + 18.15 - mmToPt(3), cY - 36));//7

    //右下
    tonbow.push(createPolyline([
        [X + W - mmToPt(3) + s, Y - H],
        [X + W - mmToPt(3), Y - H],
        [X + W - mmToPt(3), Y - H - l]
    ]));//8
    tonbow.push(createPolyline([
        [X + W, Y - H + mmToPt(3) - s],
        [X + W, Y - H + mmToPt(3)],
        [X + W + l, Y - H + mmToPt(3)]
    ]));//9

    //下辺中央
    tonbow.push(createLine(cX, Y - H - 12.15 + mmToPt(3), cX, Y - H - 36.15 + mmToPt(3)));//10
    tonbow.push(createLine(cX - 36 , Y - H - 18.15 + mmToPt(3), cX + 36, Y - H - 18.15 + mmToPt(3)));//11

    //左下
    tonbow.push(createPolyline([
        [X - l, Y - H + mmToPt(3)],
        [X, Y - H + mmToPt(3)],
        [X, Y - H + mmToPt(3) - s]
    ]));//12
    tonbow.push(createPolyline([
        [X + mmToPt(3) - s, Y - H],
        [X + mmToPt(3), Y - H],
        [X + mmToPt(3), Y - H - l]
    ]));//13

    // 左辺中央
    tonbow.push(createLine(X - 36.15 + mmToPt(3), cY, X - 12.15 + mmToPt(3), cY));//14
    tonbow.push(createLine(X - 18.15 + mmToPt(3), cY + 36, X - 18.15  + mmToPt(3), cY - 36));//15


return tonbow
};
//コピーして移動する関数
    function moveAndCopy(doc, obj, x, y) {
        var dupe = [];
        var tempArray = [];
        var ob = obj[0].duplicate();
        ob.translate(x, y);
        dupe.push(ob);
        for(var i = 0; i < obj[1].length; i++) {
            try {
                var tonbo = obj[1][i].duplicate();
                tonbo.translate(x, y);
                tempArray.push(tonbo);
            } catch(e) {}
        }
        dupe.push(tempArray);
        return dupe;
    }

(function () {
    var doc = app.activeDocument;
    var sel = doc.selection;
    if (sel.length < 1) {
        alert("最低でも1つはオブジェクトを選択してください。");
        return;
    }

    var ab = doc.artboards;
    if(ab.length < sel.length) {
        alert('アートボード数よりも多くオブジェクトを選択することはできません。');
        return;
    }
    //選択した全オブジェクトについて作業
    for(var i = 0; i < sel.length; i++){
        var obj = sel[i];
        if ( obj.typename == "GroupItem" && obj.clipped == true ) {
            var bounds = obj.pageItems[0].geometricBounds;
            var width = bounds[2] - bounds[0];
            var height = bounds[1] - bounds[3];
        } else {
            var width = obj.width;
            var height = obj.height;
        }
        var ot = [];
        var objects = [];
        ot.push(obj);
        //アートボードを選択
        var artBoard = ab[i];
        //アートボードに置ける数を横と縦で取得
        //colは横、rowは縦に置ける数
        var rect = artBoard.artboardRect;
        var abWidth = rect[2] - rect[0];
        var abHeight = rect[1] - rect[3];
        var col = Math.floor(abWidth / (width + mmToPt(6))) - 1;
        var row = Math.floor(abHeight / (height + mmToPt(6))) - 1;
        var count = (col + 1) * (row + 1) - 1;
        //トンボを生成
        ot.push(createTonbow(doc, obj));
        objects.push(ot);
        //横にコピー
        for(var j = 0; j < col; j++) {
            var delta = (j + 1) * width;
            var newt = moveAndCopy(doc, ot, delta, 0);
            objects.push(newt);
        }
        //コピーした横一列を縦にコピー
        for(var j = 0; j < row; j++) {
            var delta = (-1) * (j + 1) * height;
            for (var k = 0; k <= col; k++) {
                var newt = moveAndCopy(doc, objects[k], 0, delta);
                objects.push(newt);
            }
        }
        var LTX, LTY, RBX, RBY;
        if (obj.typename == "GroupItem" && obj.clipped == true) {
            var bounds = obj.pageItems[0].geometricBounds;
            LTX = bounds[0] + mmToPt(2);
            LTY = bounds[1] - mmToPt(2);
        } else {
            LTX = obj.position[0] + mmToPt(2);
            LTY = obj.position[1] - mmToPt(2);
        }
        
        var lastObj = objects[count][0];
        if (lastObj.typename == "GroupItem" && lastObj.clipped == true) {
            var bounds = lastObj.pageItems[0].geometricBounds;
            RBX = bounds[2] - mmToPt(2);
            RBY = bounds[3] + mmToPt(2);
        } else {
            RBX = lastObj.geometricBounds[2] - mmToPt(2);
            RBY = lastObj.geometricBounds[3] + mmToPt(2);
        }

        for(var j = 0; j <= count; j++) {
            var tonbo = objects[j][1];
            for(var k = tonbo.length - 1; k >= 0; k--) {
                try{
                    var bounds = tonbo[k].geometricBounds;
                    if((bounds[0] >= LTX && bounds[0] <= RBX && bounds[1] <= LTY && bounds[1] >= RBY) || (bounds[2] >= LTX && bounds[2] <= RBX && bounds[3] <= LTY && bounds[3] >= RBY)) {
                        tonbo[k].remove();
                    }
                } catch(e){}
            }
        }
        var allObjects = [];
        for (var j = 0; j <= count; j++) {
            allObjects.push(objects[j][0]);
            for (var k = 0; k < objects[j][1].length; k++) {
                allObjects.push(objects[j][1][k]);
            }
        }
	var allGroup = doc.groupItems.add();
	for (var j = allObjects.length - 1; j >= 0; j--) {
	    try {allObjects[j].moveToBeginning(allGroup);} catch(e){};
	}
	var gW = allGroup.width;
	var gH = allGroup.height;
	var newX = rect[0] + (abWidth - gW) / 2;
	var newY = rect[1] - (abHeight - gH) / 2;
	allGroup.position = [newX, newY];
    }
    alert("処理が完了しました。");
})();
