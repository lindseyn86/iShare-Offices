/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Environment, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import Offices from './components/Offices';
import { IOfficesProps } from './components/IOfficesProps';
import SharePointService from '../../services/SharePoint/spService';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IOfficesWebPartProps {
  webpartTitle: string;
  description: string;
  listId: string;
  region: string;
  regionfield: string;
  country: string;
  countryflag: string;
  city: string;
  leads: string;
  hrlead: string;
  itlead: string;
  officemanager: string;
  otherkeycontacts: string;
  leadstext: string;
  hrleadtext: string;
  itleadtext: string;
  officemanagertext: string;
  otherkeycontactstext: string;
  button1: string;
  button2: string;
  button3: string;
}

export default class OfficesWebPart extends BaseClientSideWebPart<IOfficesWebPartProps> {

  private webPartContext: WebPartContext;

  //list options state
  private listOptions: IPropertyPaneDropdownOption[];
  private listOptionsLoading: boolean = false;

  // field options state
  private fieldOptions: IPropertyPaneDropdownOption[];
  private fieldOptionsLoading: boolean = false;
  
  public render(): void {
    const element: React.ReactElement<IOfficesProps> = React.createElement(
      Offices,
      {
        webpartTitle: this.properties.webpartTitle,
        description: this.properties.description,
        listId: this.properties.listId,
        region: this.properties.region,
        regionfield: this.properties.regionfield,
        country: this.properties.country,
        countryflag: this.properties.countryflag,
        city: this.properties.city,
        leads: this.properties.leads,
        hrlead: this.properties.hrlead,
        itlead: this.properties.itlead,
        officemanager: this.properties.officemanager,
        otherkeycontacts: this.properties.otherkeycontacts,
        webPartContext: this.webPartContext,
        leadstext: this.properties.leadstext,
        hrleadtext: this.properties.hrleadtext,
        itleadtext: this.properties.itleadtext,
        officemanagertext: this.properties.officemanagertext,
        otherkeycontactstext: this.properties.otherkeycontactstext,
        button1: this.properties.button1,
        button2: this.properties.button2,
        button3: this.properties.button3
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.webPartContext = this.context;  // Store the context

    return super.onInit().then(() => {
      SharePointService.setup(this.context, Environment.type);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: 'Web Part Settings',
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('description', {
                  label: 'Web Part Description'
                })
              ]
            },
            {
              groupName: 'List Settings',
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: 'List',
                  options: this.listOptions,
                  disabled: this.listOptionsLoading
                }),
                PropertyPaneTextField('region', {
                  label: 'Region Text'
                }),
                PropertyPaneDropdown('regionfield', {
                  label: 'Region Title',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('country', {
                  label: 'Country',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('countryflag', {
                  label: 'Country Flag',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('city', {
                  label: 'City',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('leads', {
                  label: 'Leads',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('leadstext', {
                  label: 'Leads Text',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('hrlead', {
                  label: 'HR Lead',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('hrleadtext', {
                  label: 'HR Lead Text',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('itlead', {
                  label: 'IT Lead',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('itleadtext', {
                  label: 'IT Lead Text',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('officemanager', {
                  label: 'Office Manager',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('officemanagertext', {
                  label: 'Office Manager Text',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('otherkeycontacts', {
                  label: 'Other Key Contacts',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('otherkeycontactstext', {
                  label: 'Other Key Contacts Text',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('button1', {
                  label: 'Button 1',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('button2', {
                  label: 'Button 2',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('button3', {
                  label: 'Button 3',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getLists(): Promise<IPropertyPaneDropdownOption[]> {
    this.listOptionsLoading = true;
    this.context.propertyPane.refresh();

    return SharePointService.getLists().then(lists => {
      this.listOptionsLoading = false;
      this.context.propertyPane.refresh();
  
      return lists.value.map(list => {
        return {
          key: list.Id,
          text: list.Title
        };
      });
    });
  }

  public getFields(): Promise<any> {
    //no list selected
    if(!this.properties.listId) return Promise.resolve();
    
    this.fieldOptionsLoading = true;
    this.context.propertyPane.refresh();

    return SharePointService.getListFields(this.properties.listId).then(fields => {
      this.fieldOptionsLoading = false;
      this.context.propertyPane.refresh();
  
      return fields.value.map(field => {
        return {
          key: `${field.InternalName}+${field.Title}`,
          text: `${field.Title} (${field.TypeAsString})`
        };
      });
    });
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this.getLists().then(listOptions => {
      this.listOptions = listOptions;
      this.context.propertyPane.refresh();
    }).then(async () => {
      await this.getFields().then(fieldOptions => {
        this.fieldOptions = fieldOptions;
        this.context.propertyPane.refresh();
      });
    });
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();

    if(propertyPath === 'listId' && newValue) {
      this.properties.region = "";
      this.properties.regionfield = "";
      this.properties.country = "";
      this.properties.countryflag = "";
      this.properties.city = "";
      this.properties.leads = "";
      this.properties.hrlead = "";
      this.properties.itlead = "";
      this.properties.officemanager = "";
      this.properties.otherkeycontacts = "";
      this.properties.button1 = "";
      this.properties.button2 = "";
      this.properties.button3 = "";

      await this.getFields().then(fieldOptions => {
        this.fieldOptions = fieldOptions;
        this.context.propertyPane.refresh();
      });
    }
  }
}
