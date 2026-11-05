// WPS 功能区回调。WPS 加载 index.html 时执行，必须在全局作用域。
// 浏览器模式下这些函数不会被调用。

var WPS_Enum = { msoCTPDockPositionLeft: 0, msoCTPDockPositionRight: 2 };

// 获取当前加载项网页所在目录，用于创建同目录下的任务窗格页面。
function GetUrlPath() {
  let e = document.location.toString();
  if (e.indexOf("/") !== -1) e = decodeURI(e).substring(0, e.lastIndexOf("/"));
  return e;
}

// WPS 首次解析 ribbon.xml 时调用，保存 WPS 提供的 ribbonUI 和枚举对象。
function OnAddinLoad(ribbonUI) {
  if (typeof window.Application.ribbonUI !== "object") {
    window.Application.ribbonUI = ribbonUI;
  }
  if (typeof window.Application.Enum !== "object") {
    window.Application.Enum = WPS_Enum;
  }
  return true;
}

// 功能区按钮统一回调；当前只处理“打开 BOM 导入”按钮。
function OnAction(control) {
  if (control.Id === "btnShowMappingPane") {
    openMappingTaskPane();
  }
  return true;
}

// 所有功能区按钮默认启用。
function OnGetEnabled() {
  return true;
}

// 返回功能区按钮使用的本地图标路径。
function GetImage() {
  return "images/field-mapping.svg";
}

// 创建或切换右侧任务窗格；同一窗格 ID 会被复用，避免重复打开多个面板。
function openMappingTaskPane() {
  const app = window.Application;
  const storedId = app.PluginStorage.getItem("field_mapping_taskpane_id");
  if (!storedId) {
    const pane = app.CreateTaskPane(GetUrlPath() + "/index.html");
    app.PluginStorage.setItem("field_mapping_taskpane_id", pane.ID);
    pane.Visible = true;
  } else {
    const pane = app.GetTaskPane(storedId);
    pane.Visible = !pane.Visible;
  }
}

// 仅供 E2E 调试：在加载项宿主页上下文里执行 openMappingTaskPane()
window.__wpsFieldMappingOpenPane = openMappingTaskPane;
