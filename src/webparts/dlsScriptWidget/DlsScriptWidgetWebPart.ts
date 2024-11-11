/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable guard-for-in */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { DisplayMode } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  IPropertyPaneGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPPermission } from '@microsoft/sp-page-context'
import axios from 'axios'

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
}

export default class DlsScriptWidgetWebPart extends BaseClientSideWebPart<IDlsScriptWidgetWebPartProps> {
  private _externalContent
  private loadingIndicator: boolean = true
  private listsdisabled: boolean = true
  private widgetsdisabled: boolean = true
  private lists: IPropertyPaneDropdownOption[]
  private widgets: IPropertyPaneDropdownOption[]

  public camelToDash = (name: string) => {
    return name
      .replace(/\W+/g, '-')
      .replace(/([a-z\d])([A-Z])/g, '$1-$2')
      .toLowerCase()
  }

  constructor() {
    super();
  }

  public getDOMElementHTML(appID: string, properties: any, instanceId: string, content: any): string {
    const props = new Array<any>()
    let propAttributes: string = ''
    for (const k in properties) {
      let prop: any = properties[k]
      if (typeof prop === 'string')
        props.push({
          key: k,
          value: prop
        })
      else if (typeof prop === 'number')
        props.push({
          key: k,
          value: prop.toString()
        })
      else if (typeof prop === 'boolean')
        props.push({
          key: k,
          value: prop ? 'true' : 'false'
        })
      else if (typeof prop === 'object')
        props.push({
          key: k,
          value: JSON.stringify(prop)
        })
    }
    props.forEach((prop) => {
      propAttributes += `data-${this.camelToDash(prop.key)}="${encodeURIComponent(prop.value)}"`
    })
    return `<div id="${appID}" data-instance-id="${instanceId}" ${propAttributes}>${content}</div>`
  }

  private async getLists(): Promise<IPropertyPaneDropdownOption[]> {
    const k = new Array<IPropertyPaneDropdownOption>()
    if (this.properties.widgetsSite.length > 2) {
      this.loadingIndicator = true
      let j: any
      const surl = this.properties.widgetsSite + "/_api/web/lists"
      const response = await axios.get(surl, {
        headers: {
          accept: 'application/json;odata=verbose'
        }
      })
      j = response.data.d.results
      for (let i = 0; i < j.length; i++) {
        const hidden = Boolean(j[i].Hidden)
        const catalog = Boolean(j[i].isCatalog)
        const title = String(j[i].Title)
        const template = String(j[i].BaseTemplate)
        if (!hidden && !catalog && template === '101') {
          k.push({
            key: j[i].Id,
            text: title
          })
        }
      }
    } else  {
      this.loadingIndicator = false
    }
    return k
  }

  private async getWidgets(): Promise<IPropertyPaneDropdownOption[]> {
    const k = new Array<IPropertyPaneDropdownOption>()
    const surl = this.properties.widgetsSite + "/_api/web/lists/getbytitle('" + this.properties.widgetsList + "')/items?$select=*"
    const response = await axios.get(surl, {
      headers: {
        accept: 'application/json;odata=verbose'
      }
    })
    const j = response.data.d.results
    for (let i = 0; i < j.length; i++) {
      k.push({
        key: j[i].URL,
        text: j[i].Title
      })
    }
    return k
  }

  public render(): void {
    if (this.displayMode === DisplayMode.Read) {
      const renderDivID = "DLSWP_" + this.properties.instanceid
      const appId = "APP_" + this.properties.instanceid
      this.domElement.innerHTML = `<div id="${renderDivID}"></div>`;
      const renderDiv  = document.getElementById(renderDivID) as HTMLElement
      const content  =  this._externalContent
      renderDiv.innerHTML = this.getDOMElementHTML(appId, this.properties, this.properties.instanceid, content)
    }
  }

  protected async onInit(): Promise<void> {
    this.properties.fullcontrol = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb)
    this.properties.displayname = this.context.pageContext.user.displayName
    this.properties.loginname = this.context.pageContext.user.loginName
    this.properties.email = this.context.pageContext.user.email
    this.properties.siteadmin = this.context.pageContext.legacyPageContext.isSiteAdmin
    this.properties.instanceid = this.context.instanceId
    this.properties.siteurl = this.context.pageContext.site.absoluteUrl
    this.properties.siteid = String(this.context.pageContext.site.id)
    this.properties.weburl = this.context.pageContext.web.absoluteUrl
    this.properties.webid = String(this.context.pageContext.web.id)
    try {
      const prefix = this.properties.widget.indexOf('?') === -1 ? '?' : '&'
      const response = await fetch(`${this.properties.widget}${prefix}pnp=${new Date().getTime()}`)
      this._externalContent = await response.text()
    } catch (e) {
      this._externalContent = "Failed to load external conent."
    }
    return super.onInit()
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    switch(propertyPath) {
      case 'widgetsSite': {
        // is there a selected site url
        if (String(newValue).toLowerCase().indexOf('sharepoint') > 0) {
          // assume this is a valid site url
          this.loadingIndicator = true
          super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue)
          this.properties.widgetsList = ""
          this.listsdisabled = true
          this.context.propertyPane.refresh()
          const listOps: IPropertyPaneDropdownOption[] =  await this.getLists()
          this.lists = listOps
          this.listsdisabled = false
          this.loadingIndicator = false
          this.context.propertyPane.refresh()
        } else {
          // not a valid site
          this.loadingIndicator = false
          alert('The selected site is not a SharePoint site')
        }
        break
      }
      case 'widgetsList': {
        this.loadingIndicator = true
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue)
        this.properties.widget = ""
        this.widgetsdisabled = true
        this.context.propertyPane.refresh()
        const widgetOps: IPropertyPaneDropdownOption[] =  await this.getWidgets()
        this.widgets = widgetOps
        this.widgetsdisabled = false
        this.loadingIndicator = false
        this.context.propertyPane.refresh()
        break
      }
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const GroupFields: IPropertyPaneGroup["groupFields"] = [
      PropertyPaneTextField('widgetsSite', {
        label: 'Widget Site',
        value: this.properties.widgetsSite,
        description: "Paste Widget Site URL Here"
      }),
      PropertyPaneDropdown('list', {
        label: "Select List",
        options: this.lists,
        selectedKey: this.properties.widgetsList,
        disabled: this.listsdisabled
      }),
      PropertyPaneDropdown('widget', {
        label: 'Select Widget',
        options: this.widgets,
        selectedKey: this.properties.widget,
        disabled: this.widgetsdisabled
      })
    ]
    return {
      showLoadingIndicator: this.loadingIndicator,
      pages: [
        {
          header: {
            description: 'Web Part Settings'
          },
          groups: [
            {
              groupName: 'Group Name',
              groupFields: GroupFields
            }
          ]
        }
      ]
    };
  }
}
