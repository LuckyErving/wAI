//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI){
    if (typeof (window.Application.ribbonUI) != "object"){
		window.Application.ribbonUI = ribbonUI
    }
    
    if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
        window.Application.Enum = WPS_Enum
    }

    window.Application.PluginStorage.setItem("EnableFlag", false) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    window.Application.PluginStorage.setItem("ApiEventFlag", false) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
    return true
}

var WebNotifycount = 0;
function OnAction(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            {
                const doc = window.Application.ActiveDocument
                if (!doc) {
                    alert("当前没有打开任何文档")
                    return
                }
                alert(doc.Name)
            }
            break;
        case "btnIsEnbable":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                window.Application.PluginStorage.setItem("EnableFlag", !bFlag)
                
                //通知wps刷新以下几个按饰的状态
                window.Application.ribbonUI.InvalidateControl("btnIsEnbable")
                window.Application.ribbonUI.InvalidateControl("btnShowDialog") 
                window.Application.ribbonUI.InvalidateControl("btnShowTaskPane") 
                //window.Application.ribbonUI.Invalidate(); 这行代码打开则是刷新所有的按钮状态
                break
            }
        case "btnShowDialog":
            window.Application.ShowDialog(GetUrlPath() + "/ui/dialog.html", "这是一个对话框网页", 400 * window.devicePixelRatio, 400 * window.devicePixelRatio, false)
            break
        case "btnShowTaskPane":
            {
                let tsId = window.Application.PluginStorage.getItem("taskpane_id")
                if (!tsId) {
                    let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
                    let id = tskpane.ID
                    window.Application.PluginStorage.setItem("taskpane_id", id)
                    tskpane.Visible = true
                } else {
                    let tskpane = window.Application.GetTaskPane(tsId)
                    tskpane.Visible = !tskpane.Visible
                }
            }
            break
        case "btnApiEvent":
            {
                let bFlag = window.Application.PluginStorage.getItem("ApiEventFlag")
                let bRegister = bFlag ? false : true
                window.Application.PluginStorage.setItem("ApiEventFlag", bRegister)
                if (bRegister){
                    window.Application.ApiEvent.AddApiEventListener('DocumentNew', OnNewDocumentApiEvent)
                }
                else{
                    window.Application.ApiEvent.RemoveApiEventListener('DocumentNew', OnNewDocumentApiEvent)
                }

                window.Application.ribbonUI.InvalidateControl("btnApiEvent") 
            }
            break
        default:
            // 新增AI按钮处理
            if (eleId === "btnAIRewrite") {
                handleAIAction("rewrite");
            } else if (eleId === "btnAISummary") {
                handleAIAction("summary");
            } else if (eleId === "btnWithdrawAI") {
                withdrawLastAIInsert();
            }
            break;
    }
    return true;
}

// 用于记录上次AI插入的起止位置和原内容
let lastAIInsert = { start: null, end: null, originalText: null };
let aiInsertInProgress = false; // 标记AI插入是否进行中

// AI处理函数
async function handleAIAction(type) {
    const doc = window.Application.ActiveDocument;
    if (!doc) {
        alert("当前没有打开任何文档");
        return;
    }
    let sel = window.Application.Selection;
    if (!sel || !sel.Range || sel.Range.Text === "") {
        alert("请先选中需要处理的文本");
        return;
    }
    let selectedText = sel.Range.Text;
    let prompt = "";
    if (type === "rewrite") {
        prompt = `请对以下内容进行续写和润色，只返回润色后的正文内容，不要输出任何解释或格式说明：${selectedText}`;
    } else if (type === "summary") {
        prompt = `请对以下内容进行总结提炼，只返回总结后的正文内容，不要输出任何解释或格式说明：${selectedText}`;
    }
    // 调用AI接口
    try {
        aiInsertInProgress = true;
        updateWithdrawAIButton();
        const aiResult = await callAIAPI(prompt);
        let range = sel.Range;
        // 记录插入前的起点和原内容
        let insertStart = range.Start;
        let insertEnd = range.End;
        let originalText = range.Text;
        range.Text = ""; // 先清空选区
        // 打字机效果插入
        for (let i = 0; i < aiResult.length; i++) {
            await new Promise(resolve => setTimeout(resolve, 400)); // 100ms/字
            range.Text += aiResult[i];
            range.Start = range.Start + 1;
            range.End = range.Start;
            range.Font.Color = 255; // 红色
        }
        // 插入一个回车，防止与后文接触
        range.Text += "\r";
        range.Start = range.Start + 1;
        range.End = range.Start;
        // 记录插入的起止位置和原内容
        lastAIInsert.start = insertStart;
        lastAIInsert.end = insertStart + aiResult.length + 1; // +1 for the new line
        lastAIInsert.originalText = originalText;
    } catch (e) {
        alert("AI接口调用失败：" + e.message);
    } finally {
        aiInsertInProgress = false;
        updateWithdrawAIButton();
    }
}

// 控制撤回按钮可用性
function updateWithdrawAIButton() {
    if (window.Application && window.Application.ribbonUI) {
        window.Application.ribbonUI.InvalidateControl("btnWithdrawAI");
    }
}

// 撤回上次AI插入内容并恢复原内容
function withdrawLastAIInsert() {
    const doc = window.Application.ActiveDocument;
    if (!doc) {
        alert("当前没有打开任何文档");
        return;
    }
    if (lastAIInsert.start === null || lastAIInsert.end === null) {
        alert("没有可撤回的AI插入内容");
        return;
    }
    let range = doc.Range(lastAIInsert.start, lastAIInsert.end);
    range.Text = lastAIInsert.originalText || "";
    // 恢复选中原文字
    let sel = window.Application.Selection;
    sel.SetRange(lastAIInsert.start, lastAIInsert.start + (lastAIInsert.originalText ? lastAIInsert.originalText.length : 0));
    // 撤回后清空记录
    lastAIInsert.start = null;
    lastAIInsert.end = null;
    lastAIInsert.originalText = null;
    //alert("已撤回上次AI插入内容，并恢复原文字");
}

// 示例AI接口调用（需替换为你自己的API KEY和接口地址）
async function callAIAPI(prompt) {
    const apiKey = "sk-wgR14x8Ec0njyIfT22Ab2554662a494d8d18807b57200686"; // 替换为你的API KEY
    // const apiUrl = "https://api.openai.com/v1/chat/completions";  // https://free.v36.cm
    const apiUrl = "https://free.v36.cm/v1/chat/completions"; // 替换为你的API URL
    const response = await fetch(apiUrl, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            model: "gpt-3.5-turbo",
            messages: [{ role: "user", content: prompt }],
            max_tokens: 512
        })
    });
    if (!response.ok) throw new Error("AI接口请求失败");
    const data = await response.json();
    return data.choices[0].message.content.trim();
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg"
        case "btnShowDialog":
            return "images/2.svg"
        case "btnShowTaskPane":
            return "images/3.svg"
        default:
            ;
    }
    return "images/newFromTemp.svg"
}

function OnGetEnabled(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return true
        case "btnShowDialog":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                return bFlag
            }
        case "btnShowTaskPane":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                return bFlag
            }
        case "btnWithdrawAI":
            return !aiInsertInProgress && lastAIInsert.start !== null && lastAIInsert.end !== null;
        default:
            break
    }
    return true
}

function OnGetVisible(control){
    return true
}

function OnGetLabel(control){
    const eleId = control.Id
    switch (eleId) {
        case "btnIsEnbable":
        {
            let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
            return bFlag ?  "按钮Disable" : "按钮Enable"
        }
        case "btnApiEvent":
        {
            let bFlag = window.Application.PluginStorage.getItem("ApiEventFlag")
            return bFlag ? "清除新建文件事件" : "注册新建文件事件"
        }    
    }
    return ""
}

function OnNewDocumentApiEvent(doc){
    alert("新建文件事件响应，取文件名: " + doc.Name)
}
