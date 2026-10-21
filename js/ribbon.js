// WPS 功能区回调。WPS 加载 index.html 时执行，必须在全局作用域。
// 浏览器模式下这些函数不会被调用。

var WPS_Enum = { msoCTPDockPositionLeft: 0, msoCTPDockPositionRight: 2 };

function GetUrlPath() {
  let e = document.location.toString();
  if (e.indexOf("/") !== -1) e = decodeURI(e).substring(0, e.lastIndexOf("/"));
  return e;
}

function OnAddinLoad(ribbonUI) {
  if (typeof window.Application.ribbonUI !== "object") {
    window.Application.ribbonUI = ribbonUI;
  }
  if (typeof window.Application.Enum !== "object") {
    window.Application.Enum = WPS_Enum;
  }
  return true;
}

function OnAction(control) {
  if (control.Id === "btnShowMappingPane") {
    openMappingTaskPane();
  }
  return true;
}

function OnGetEnabled() {
  return true;
}

function GetImage() {
  return "images/field-mapping.svg";
}

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
