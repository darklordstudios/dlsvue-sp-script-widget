/* eslint-disable dot-notation */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable guard-for-in */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { DisplayMode } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration, PropertyPaneTextField, IPropertyPaneDropdownOption, PropertyPaneDropdown, IPropertyPaneGroup } from "@microsoft/sp-property-pane";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { SPPermission } from "@microsoft/sp-page-context";
import axios from "axios";

export interface IDlsScriptWidgetWebPartProps {
  description?: string;
  widget: string;
  widgetsSite: string; // URL where list of widgets is
  widgetsList: string; // Name of widgest list
  title?: string;
  editmode: boolean;
  siteurl: string;
  siteid: string;
  weburl: string;
  webid: string;
  instanceid: string;
  fullcontrol: boolean;
  displayname: string;
  loginname: string;
  email: string;
  siteadmin: boolean;
  renderelement: string;
}

export default class DlsScriptWidgetWebPart extends BaseClientSideWebPart<IDlsScriptWidgetWebPartProps> {
  private _externalContent;
  private loadingIndicator: boolean = false;
  private listsdisabled: boolean = true;
  private widgetsdisabled: boolean = true;
  private lists: IPropertyPaneDropdownOption[];
  private widgets: IPropertyPaneDropdownOption[];
  public widgetinfos = new Array<any>()
  private msg: string = "Welcome";

  constructor() {
    super();
  }

  private async getLists(): Promise<IPropertyPaneDropdownOption[]> {
    const k = new Array<IPropertyPaneDropdownOption>();
    if (this.properties.widgetsSite.length > 2) {
      this.loadingIndicator = true;
      let j: any;
      const surl = this.properties.widgetsSite + "/_api/web/lists";
      const response = await axios.get(surl, {
        headers: {
          accept: "application/json;odata=verbose",
        },
      });
      j = response.data.d.results;
      for (let i = 0; i < j.length; i++) {
        const hidden = Boolean(j[i].Hidden);
        const catalog = Boolean(j[i].isCatalog);
        const title = String(j[i].Title);
        const template = String(j[i].BaseTemplate);
        if (!hidden && !catalog && template === "100") {
          k.push({
            key: title,
            text: title,
          });
        }
      }
    } else {
      this.loadingIndicator = false;
    }
    return k;
  }

  private async getWidgets(): Promise<IPropertyPaneDropdownOption[]> {
    const k = new Array<IPropertyPaneDropdownOption>();
    const surl = this.properties.widgetsSite + "/_api/web/lists/getbytitle('" + this.properties.widgetsList + "')/items?$select=*";
    const response = await axios.get(surl, {
      headers: {
        accept: "application/json;odata=verbose",
      },
    });
    const j = response.data.d.results;
    for (let i = 0; i < j.length; i++) {
      const status = j[i].Status
      if (status === 'Production') {
        k.push({
          index: i,
          key: j[i].URL,
          text: j[i].Title
        })
        this.widgetinfos.push({
          index: i,
          url: j[i].URL,
          title: j[i].Title,
          status: j[i].Status,
          element: j[i].Element,
          details: j[i].Details
        })
      } else {
        if (this.properties.siteadmin === true) {
          k.push({
            index: i,
            key: j[i].URL,
            text: j[i].Title
          })
          this.widgetinfos.push({
            index: i,
            url: j[i].URL,
            title: j[i].Title,
            status: j[i].Status,
            element: j[i].Element,
            details: j[i].Details
          })
        }
      }
    }
    return k;
  }

  private evalScript(elem: { text: any; textContent: any; innerHTML: any; attributes: string | any[] }) {
    const data = elem.text || elem.textContent || elem.innerHTML || "";
    const headtag = document.getElementsByClassName("head")[0] || document.documentElement;
    const scriptTag = document.createElement("script");
    for (let i = 0; i < elem.attributes.length; i++) {
      const attr = elem.attributes[i];
      if (attr.name.toLowerCase() === "onload") continue;
      scriptTag.setAttribute(attr.name, attr.value);
    }
    scriptTag.type = scriptTag.src && scriptTag.src.length > 0 ? "pnp" : "text/javascript";
    scriptTag.setAttribute("pnpname", this.properties.instanceid);
    try {
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      scriptTag.text = data;
    }
    headtag.insertBefore(scriptTag, headtag.firstChild);
  }

  private async executeScript(element: HTMLElement) {
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTags = headTag.getElementsByTagName("script");
    for (let i = 0; i < scriptTags.length; i++) {
      const scriptTag = scriptTags[i];
      if (scriptTag.hasAttribute("pnpname") && scriptTag.attributes["pnpname"].value === this.properties.instanceid) {
        headTag.removeChild(scriptTag);
      }
    }

    (<any>window).ScriptGlobal = {};

    const scripts = new Array<any>();
    const children_nodes = element.getElementsByTagName("script");

    for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (!child.type || child.type.toLowerCase() === "text/javascript" || child.type.toLowerCase() === "module") {
        scripts.push(child);
      }
    }

    const urls = new Array<any>();
    const onloads = new Array<any>();
    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
        onloads.push(scriptTag.onload);
      }
    }

    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for (let i = 0; i < urls.length; i++) {
      try {
        let scriptUrl = urls[i];
        const prefix = scriptUrl.indexOf("?") === -1 ? "?" : "&";
        scriptUrl += prefix + "pnp=" + new Date().getTime();
        await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
      } catch (e) {
        if (console.error) {
          console.error(e);
        }
      }
    }

    if (oldamd) {
      window["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) {
        scriptTag.parentNode.removeChild(scriptTag);
      }
      this.evalScript(scripts[i]);
    }

    for (let i = 0; onloads[i]; i++) {
      onloads[i]();
    }
  }

  protected async onInit(): Promise<void> {
    this.properties.fullcontrol = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    this.properties.displayname = this.context.pageContext.user.displayName;
    this.properties.loginname = this.context.pageContext.user.loginName;
    this.properties.email = this.context.pageContext.user.email;
    this.properties.siteadmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
    this.properties.instanceid = this.context.instanceId;
    this.properties.siteurl = this.context.pageContext.site.absoluteUrl;
    this.properties.siteid = String(this.context.pageContext.site.id);
    this.properties.weburl = this.context.pageContext.web.absoluteUrl;
    this.properties.webid = String(this.context.pageContext.web.id);
    try {
      const prefix = this.properties.widget.indexOf("?") === -1 ? "?" : "&";
      const response = await fetch(`${this.properties.widget}${prefix}pnp=${new Date().getTime()}`);
      this._externalContent = await response.text();
    } catch (e) {
      this._externalContent = "Failed to load external conent.";
    }
    globalThis.spcontext = this.context;
    globalThis.hasFullControl = this.properties.fullcontrol;
    return super.onInit();
  }

  public async render(): Promise<void> {
    if (this.displayMode === DisplayMode.Read) {
      const renderDivID = "DLSWP_" + this.properties.instanceid;
      const appId = "APP_" + this.properties.instanceid;
      this.domElement.innerHTML = `<div id="${renderDivID}"></div>`;
      const renderDiv = document.getElementById(renderDivID) as HTMLElement;
      const content = this._externalContent;
      // renderDiv.innerHTML = this.getDOMElementHTML(appId, this.properties, this.properties.instanceid, content)
      renderDiv.innerHTML = `<div id="${appId}">${content}</div>`;
      await this.executeScript(this.domElement);
    }
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    switch (propertyPath) {
      case "widgetsSite": {
        // is there a selected site url
        if (String(newValue).toLowerCase().indexOf("sharepoint") > 0) {
          // assume this is a valid site url
          this.loadingIndicator = true;
          super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
          this.properties.widgetsList = "";
          this.listsdisabled = true;
          this.context.propertyPane.refresh();
          const listOps: IPropertyPaneDropdownOption[] = await this.getLists();
          this.lists = listOps;
          this.listsdisabled = false;
          this.loadingIndicator = false;
          this.context.propertyPane.refresh();
        } else {
          // not a valid site
          this.loadingIndicator = false;
          alert("The selected site is not a SharePoint site");
        }
        break;
      }
      case "widgetsList": {
        this.loadingIndicator = true;
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        this.properties.widget = "";
        this.widgetsdisabled = true;
        this.context.propertyPane.refresh();
        const widgetOps: IPropertyPaneDropdownOption[] = await this.getWidgets();
        this.widgets = widgetOps;
        this.widgetsdisabled = false;
        this.loadingIndicator = false;
        this.context.propertyPane.refresh();
        break;
      }
    }
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.context.propertyPane.refresh();
    if (this.properties.widgetsSite === "") {
      this.loadingIndicator = false;
      this.msg = "Paste URL to Widget Site";
    } else {
      this.loadingIndicator = true;
      const listOps: IPropertyPaneDropdownOption[] = await this.getLists();
      this.lists = listOps;
      this.listsdisabled = false;
      this.loadingIndicator = false;
      this.context.propertyPane.refresh();
      if (this.properties.widget !== "") {
        // get the widgets
        this.loadingIndicator = true;
        const widgetOps: IPropertyPaneDropdownOption[] = await this.getWidgets();
        this.widgets = widgetOps;
        this.widgetsdisabled = false;
        this.loadingIndicator = false;
        this.context.propertyPane.refresh();
      }
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const GroupFields: IPropertyPaneGroup["groupFields"] = [
      PropertyPaneTextField("msg", {
        label: "Message",
        value: this.msg,
        disabled: true,
      }),
      PropertyPaneTextField("widgetsSite", {
        label: "Widget Site",
        value: this.properties.widgetsSite,
        description: "Paste Widget Site URL Here",
      }),
      PropertyPaneDropdown("widgetsList", {
        label: "Select List",
        options: this.lists,
        selectedKey: this.properties.widgetsList,
        disabled: this.listsdisabled,
      }),
      PropertyPaneDropdown("widget", {
        label: "Select Widget",
        options: this.widgets,
        selectedKey: this.properties.widget,
        disabled: this.widgetsdisabled,
      }),
    ];
    return {
      showLoadingIndicator: this.loadingIndicator,
      pages: [
        {
          header: {
            description: "Web Part Settings",
          },
          groups: [
            {
              groupName: "Group Name",
              groupFields: GroupFields,
            },
          ],
        },
      ],
    };
  }
}
