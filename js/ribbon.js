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
        case "btnWebNotify":
            {
                let currentTime = new Date()
                let timeStr = currentTime.getHours() + ':' + currentTime.getMinutes() + ":" + currentTime.getSeconds()
                window.Application.OAAssist.WebNotify("这行内容由wps加载项主动送达给业务系统，可以任意自定义, 比如时间值:" + timeStr + "，次数：" + (++WebNotifycount), true)
            }
            break
        default:
            // 新增AI按钮处理
            if (eleId === "btnAIRewrite") {
                handleAIAction("rewrite");
            } else if (eleId === "btnAISummary") {
                handleAIAction("summary");
            }
            break;
    }
    return true;
}

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
        prompt = `请对以下内容进行续写和润色：${selectedText}`;
    } else if (type === "summary") {
        prompt = `请对以下内容进行总结提炼：${selectedText}`;
    }
    // 调用AI接口
    try {
        const aiResult = await callAIAPI(prompt);
        // 插入AI结果到光标处
        sel.Range.Text = aiResult;
        alert("AI处理完成，结果已插入文档");
    } catch (e) {
        alert("AI接口调用失败：" + e.message);
    }
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
            break
        case "btnShowDialog":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        case "btnShowTaskPane":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
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
            break
        }
        case "btnApiEvent":
        {
            let bFlag = window.Application.PluginStorage.getItem("ApiEventFlag")
            return bFlag ? "清除新建文件事件" : "注册新建文件事件"
            break
        }    
    }
    return ""
}

function OnNewDocumentApiEvent(doc){
    alert("新建文件事件响应，取文件名: " + doc.Name)
}
