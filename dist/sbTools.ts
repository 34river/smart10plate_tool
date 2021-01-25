function onOpen() {
    SlidesApp.getUi()
        .createMenu('SmartTemplate作成支援')
        .addItem('サイドバー', 'showSidebar')
        .addToUi();
}

function showSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('sbTool')
        .setTitle('SmartTemplate作成支援')
        .setWidth(300);
    SlidesApp.getUi()
        .showSidebar(html);
}


// プレースホルダーの説明を切り替える
// status: ture 表示　false 非表示
function phDescVisible(status: boolean) {
    //マスターを取得
    const masters = SlidesApp.getActivePresentation().getMasters();
    console.log(`master数=${masters.length}`);

    const layouts = SlidesApp.getActivePresentation().getLayouts();
    console.log(`layout数=${layouts.length}`);

    for (const layout of layouts) {

        //マスター配下のShapeを取得  
        const shapes = layout.getShapes();
        console.log(`shpes数=${shapes.length}`);
        for (var i = 0; i < shapes.length; ++i) {

            const msg = shapes[i].getTitle();
            if (msg.startsWith('PH_')) {
                //表示
                let desc = shapes[i].getDescription();
                if (status === true) {
                    if (desc != "") {
                        const textObj = shapes[i].getText();
                        textObj.setText(desc);
                        /* 半角スペースを設定すれば段落情報が残り続けるのでいちいち設定しなくてもよい。
                        textObj.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
                        textObj.getTextStyle().setFontSize(20)
                                              .setBold(true)
                                              .setForegroundColor('#FF0000');
                        */
                    }
                } else {
                    //非表示 半角スペースを設定しておけば段落情報は残り続ける。
                    shapes[i].getText().setText(" ");
                }
                if (msg.startsWith('PH_TYPE01')) {
                    const border = shapes[i].getBorder();
                    const fill = shapes[i].getFill();
                    if (status === false) {
                        //OFF
                        fill.setTransparent();
                        border.setTransparent();
                    }
                } else if (msg.startsWith('PH_TYPE02')) {
                    const border = shapes[i].getBorder();
                    const fill = shapes[i].getFill();
                    if (status === false) {
                        //OFF
                        fill.setTransparent();
                        border.setTransparent();
                    } else {
                        //ON
                        fill.setSolidFill("#FFFFFF");
                        border.getLineFill().setSolidFill("#000000");
                    }
                }
            }
        }
    }
}


// プレースホルダーの説明を切り替える
// status: ture 表示　false 非表示
function phDescDebugVisible(status: boolean) {
    //マスターを取得
    const masters = SlidesApp.getActivePresentation().getMasters();
    const layouts = SlidesApp.getActivePresentation().getLayouts();

    for (const layout of layouts) {

        //マスター配下のShapeを取得  
        const shapes = layout.getShapes();
        for (var i = 0; i < shapes.length; ++i) {
            const msg = shapes[i].getTitle();
            if (msg.startsWith('PH_')) {
                //説明文をShapeのテキストにセットする。
                let desc = shapes[i].getDescription();
                if (status === true) {
                    if (desc != "") {
                        const textObj = shapes[i].getText();
                        textObj.setText(desc);
                        /* 半角スペースを設定すれば段落情報が残り続けるのでいちいち設定しなくてもよい。
                        textObj.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
                        textObj.getTextStyle().setFontSize(20)
                                              .setBold(true)
                                              .setForegroundColor('#FF0000');
                        */
                    }
                } else {
                    //非表示 半角スペースを設定しておけば段落情報は残り続ける。
                    shapes[i].getText().setText(" ");
                }
                if (msg.startsWith('PH_TYPE01')) {
                    if (status === true) {
                        //分かりやすくするため、枠線を表示しておく
                        // ダッシュで赤を指定する。
                        shapes[i].getBorder().setWeight(2)
                            .setDashStyle(SlidesApp.DashStyle.DASH)
                            .getLineFill().setSolidFill("#FF0000");
                    }

                }
                if (msg.startsWith('PH_TYPE02')) {
                    const border = shapes[i].getBorder();
                    const fill = shapes[i].getFill();
                    if (status === false) {
                        //OFF
                        fill.setTransparent();
                        border.setTransparent();
                    } else {
                        //ON
                        fill.setSolidFill("#FFFFFF");
                        border.getLineFill().setSolidFill("#000000");
                    }
                }
            }
        }
    }
}


function phDescShapeSet() {
    //指定したShapeをDesc化する。
    const retSelection = shapeCheckGet();
    if (typeof (retSelection) == "string") {
        SlidesApp.getUi().alert(retSelection);
        return;
    }
    if (retSelection.getParentPage().getPageType() != SlidesApp.PageType.LAYOUT) {
        SlidesApp.getUi().alert("ガイダンスはマスターのレイアウトに配置してください");
        return;
    }
    //文字が設定されているか？
    const shapeText = shapeTextGet(retSelection);
    if (shapeText.ret == false) {
        SlidesApp.getUi().alert(shapeText.text);
        return;
    }
    //すでにガイド設定されているか？
    if (retSelection.getTitle() == "PH_TYPE02") {
        SlidesApp.getUi().alert("PH_TYPE02として設定されています");
        return;
    }
    retSelection.setTitle("PH_TYPE01");
    retSelection.setDescription(shapeText.text);

    SlidesApp.getUi().alert("PH_TYPE01として登録しました。");
}
function phDescShapeSet2() {
    //指定したShapeをDesc化する。
    const retSelection = shapeCheckGet();
    if (typeof (retSelection) == "string") {
        SlidesApp.getUi().alert(retSelection);
        return;
    }
    if (retSelection.getParentPage().getPageType() != SlidesApp.PageType.LAYOUT) {
        SlidesApp.getUi().alert("ガイダンスはマスターのレイアウトに配置してください");
        return;
    }
    //文字が設定されているか？
    const shapeText = shapeTextGet(retSelection);
    if (shapeText.ret == false) {
        SlidesApp.getUi().alert(shapeText.text);
        return;
    }
    //すでにガイド設定されているか？
    if (retSelection.getTitle() == "PH_TYPE01") {
        SlidesApp.getUi().alert("PH_TYPE01として設定されています");
        return;
    }
    retSelection.setTitle("PH_TYPE02");
    retSelection.setDescription(shapeText.text);

    SlidesApp.getUi().alert("PH_TYPE02として登録しました。");
}

function shapeCheckGet() {
    //編集対象の文が選択されているか？
    const selection = SlidesApp.getActivePresentation().getSelection();
    if (selection == null) {
        const msg: string = "テキストボックスを1つ選択して下さい。";
        return msg;
    }
    const selectionType = selection.getSelectionType();
    if (selectionType == SlidesApp.SelectionType.TEXT) {
        //テキストが選択されている場合は一つ上をチェックする。
        const pageElements = selection.getPageElementRange().getPageElements();
        pageElements[0].select();
        return pageElements[0];
    }
    else if (selectionType != SlidesApp.SelectionType.PAGE_ELEMENT) {
        const msg: string = "テキストボックスを選択して下さい";
        return msg;
    } else {

        const pageElements = selection.getPageElementRange().getPageElements();
        if (pageElements.length != 1) {
            const msg: string = "テキストボックスを1つだけ選択して下さい";
            return msg;
        }
        const PageElementType = pageElements[0].getPageElementType();
        if (PageElementType != SlidesApp.PageElementType.SHAPE) {
            const msg: string = "テキストボックスを選択して下さい";
            return msg;
        }
        //正しく取得された時はobjを返却
        return pageElements[0];
    }
}
function shapeTextGet(objShapeShape) {
    //テキスト内容を取得
    const textRanges = objShapeShape.asShape().getText().getRuns();
    if (textRanges.length != 1) {
        return { ret: false, text: "修飾を複数利用しているテキストには対応していません。" };
    }
    const textRange = textRanges[0];
    const text = objShapeShape.asShape().getText().asString();
    if (text == null) {
        return { ret: false, text: "テキストが入力されていません。" };
    }
    return { ret: true, text: text };
}
function phDescShapeZOrder(mode: string) {
    if (mode != "sendBackward" && mode != "bringToFront") {
        return false;
    }
    const masters = SlidesApp.getActivePresentation().getMasters();
    const layouts = SlidesApp.getActivePresentation().getLayouts();


    for (const layout of layouts) {

        //マスター配下のShapeを取得  
        const shapes = layout.getShapes();
        const elements = layout.getPageElements();

        for (var i = 0; i < elements.length; ++i) {
            const msg = elements[i].getTitle();
            if (msg.startsWith('PH_')) {
                //選択
                elements[i].select(false);
            }
        }

        const selection = SlidesApp.getActivePresentation().getSelection();
        const pageElements = selection.getPageElementRange().getPageElements();
        const shapesGroup = pageElements.length >= 2 ? layout.group(pageElements) : undefined;
        try {
            switch (mode) {
                case 'bringToFront':
                    shapesGroup ? shapesGroup.bringToFront() : pageElements[0].bringToFront();
                    break;
                case 'sendBackward':
                    shapesGroup ? shapesGroup.sendBackward() : pageElements[0].sendBackward();
                    break;
            }
        }
        catch (e) {
            return false;
        }
        if (shapesGroup) {
            shapesGroup.ungroup();
        }
    }

}
