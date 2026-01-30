import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IOfficesProps {
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
  webPartContext: WebPartContext;
  leadstext: string;
  hrleadtext: string;
  itleadtext: string;
  officemanagertext: string;
  otherkeycontactstext: string;
  button1: string;
  button2: string;
  button3: string;
}