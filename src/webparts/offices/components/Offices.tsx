/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-unused-expressions */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import styles from './Offices.module.scss';
import type { IOfficesProps } from './IOfficesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IOfficesState } from './IOfficesState';
import SharePointService from '../../../services/SharePoint/spService';
import { PersonaSize } from '@fluentui/react';
import SPFxPeopleCard from "./SPFxPeopleCard/SPFxPeopleCard";
import { Icon } from "@fluentui/react";

export default class Offices extends React.Component<IOfficesProps, IOfficesState> {
  constructor(props: IOfficesProps) {
    super(props);

    // bind methods
    this.getItems = this.getItems.bind(this);

    // set initial state
    this.state = {
      items: [],
      loading: false,
      // sortBy: this.props.country ? this.props.country.split('+')[0] : this.props.city ? this.props.city.split('+')[0] : '',
      // sortOrder: 'asc',
      error: 'null',
    }
  }

  public render(): React.ReactElement<IOfficesProps> {
    const {
      webpartTitle,
      description,
      webPartContext,
      country,
      countryflag,
      city,
      leads,
      hrlead,
      itlead,
      officemanager,
      otherkeycontacts,
      leadstext,
      hrleadtext,
      itleadtext,
      officemanagertext,
      otherkeycontactstext,
      button1,
      button2,
      button3
    } = this.props;

    return (
      <section className={`${styles.offices}`}>
        {webpartTitle && <h2 className={styles.webparttitle}>{webpartTitle}</h2>}
        {description && <p className={styles.description}>{escape(description)}</p>}

        <div className={styles.officescontainer}>
          {this.state.items.length > 0 && this.state.items.map((item,index) =>
            <div className={styles.office} key={index}>
              <h2 className={styles.header} style={{ backgroundImage: item[countryflag.split('+')[0]] ? `url(${item[countryflag.split('+')[0]]})` : '' }}>
                <span className={styles.country}>{item[country.split('+')[0]]}</span>
                <span className={styles.city}>{item[city.split('+')[0]]}</span>
              </h2>
              <div className={styles.contacts}>

                {leadstext && item[leadstext.split('+')[0]] && item[leadstext.split('+')[0]].split(';').length > 0 ?
                  <div className={item[leadstext.split('+')[0]].split(';').length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{leads.split('+')[1]}</h3>
                    <div className={styles.employeeList}>
                      {item[leadstext.split('+')[0]].split(';').map((lead: any,index: React.Key | null | undefined) => 
                        <div key={index} className={item[leadstext.split('+')[0]].split(';').length === 1 ? styles.employee100 : styles.employee50}>
                          <div>{lead.split(' - ')[0]}</div>
                          {lead.split(' - ').length > 1 && <div><a href={`mailto:${lead.split(' - ')[1]}`}>{lead.split(' - ')[1]}</a></div>}
                        </div>
                      )}
                    </div>
                  </div>
                :
                leads && item[leads.split('+')[0]] && item[leads.split('+')[0]].length > 0 &&
                  <div className={item[leads.split('+')[0]].length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{leads.split('+')[1]}</h3>
                    <div className={styles.employeeList}>

                      {item[leads.split('+')[0]].map((lead: { Email: string; LastName: string; FirstName: string; JobTitle: string; EMail: string; },index: React.Key | null | undefined) => 
                      
                        <SPFxPeopleCard
                          email={lead.EMail}
                          primaryText={`${lead.LastName}, ${lead.FirstName}`}
                          secondaryText={lead.JobTitle}
                          serviceScope={webPartContext.serviceScope}
                          size={PersonaSize.size48}
                          key={`peopleCard-${lead.EMail}`}
                          class={item[leads.split('+')[0]].length === 1 ? styles.employee100 : styles.employee50}
                        />

                      )}

                    </div>
                  </div>
                }

                {hrleadtext && item[hrleadtext.split('+')[0]] && item[hrleadtext.split('+')[0]].split(';').length > 0 ?
                  <div className={item[hrleadtext.split('+')[0]].split(';').length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{hrlead.split('+')[1]}</h3>
                    <div className={styles.employeeList}>
                      {item[hrleadtext.split('+')[0]].split(';').map((lead: any,index: React.Key | null | undefined) => 
                        <div key={index} className={item[hrleadtext.split('+')[0]].split(';').length === 1 ? styles.employee100 : styles.employee50}>
                          <div>{lead.split(' - ')[0]}</div>
                          {lead.split(' - ').length > 1 && <div><a href={`mailto:${lead.split(' - ')[1]}`}>{lead.split(' - ')[1]}</a></div>}
                        </div>
                      )}
                    </div>
                  </div>
                :
                hrlead && item[hrlead.split('+')[0]] && item[hrlead.split('+')[0]].length > 0 &&
                  <div className={item[hrlead.split('+')[0]].length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{hrlead.split('+')[1]}</h3>
                    <div className={styles.employeeList}>

                      {item[hrlead.split('+')[0]].map((hrlead1: { Email: string; LastName: string; FirstName: string; JobTitle: string; EMail: string; },index: React.Key | null | undefined) => 
                      
                        <SPFxPeopleCard
                          email={hrlead1.EMail}
                          primaryText={`${hrlead1.LastName}, ${hrlead1.FirstName}`}
                          secondaryText={hrlead1.JobTitle}
                          serviceScope={webPartContext.serviceScope}
                          size={PersonaSize.size48}
                          key={`peopleCard-${hrlead1.EMail}`}
                          class={item[hrlead.split('+')[0]].length === 1 ? styles.employee100 : styles.employee50}
                        />

                      )}

                    </div>
                  </div>
                }

                {itleadtext && item[itleadtext.split('+')[0]] && item[itleadtext.split('+')[0]].split(';').length > 0 ?
                  <div className={item[itleadtext.split('+')[0]].split(';').length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{itlead.split('+')[1]}</h3>
                    <div className={styles.employeeList}>
                      {item[itleadtext.split('+')[0]].split(';').map((lead: any,index: React.Key | null | undefined) => 
                        <div key={index} className={item[itleadtext.split('+')[0]].split(';').length === 1 ? styles.employee100 : styles.employee50}>
                          <div>{lead.split(' - ')[0]}</div>
                          {lead.split(' - ').length > 1 && <div><a href={`mailto:${lead.split(' - ')[1]}`}>{lead.split(' - ')[1]}</a></div>}
                        </div>
                      )}
                    </div>
                  </div>
                :
                itlead && item[itlead.split('+')[0]] && item[itlead.split('+')[0]].length > 0 &&
                  <div className={item[itlead.split('+')[0]].length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{itlead.split('+')[1]}</h3>
                    <div className={styles.employeeList}>

                      {item[itlead.split('+')[0]].map((itlead1: { Email: string; LastName: string; FirstName: string; JobTitle: string; EMail: string; },index: React.Key | null | undefined) => 
                      
                        <SPFxPeopleCard
                          email={itlead1.EMail}
                          primaryText={`${itlead1.LastName}, ${itlead1.FirstName}`}
                          secondaryText={itlead1.JobTitle}
                          serviceScope={webPartContext.serviceScope}
                          size={PersonaSize.size48}
                          key={`peopleCard-${itlead1.EMail}`}
                          class={item[itlead.split('+')[0]].length === 1 ? styles.employee100 : styles.employee50}
                        />

                      )}

                    </div>
                  </div>
                }
                
                {officemanagertext && item[officemanagertext.split('+')[0]] && item[officemanagertext.split('+')[0]].split(';').length > 0 ?
                  <div className={item[officemanagertext.split('+')[0]].split(';').length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{officemanager.split('+')[1]}</h3>
                    <div className={styles.employeeList}>
                      {item[officemanagertext.split('+')[0]].split(';').map((lead: any,index: React.Key | null | undefined) => 
                        <div key={index} className={item[officemanagertext.split('+')[0]].split(';').length === 1 ? styles.employee100 : styles.employee50}>
                          <div>{lead.split(' - ')[0]}</div>
                          {lead.split(' - ').length > 1 && <div><a href={`mailto:${lead.split(' - ')[1]}`}>{lead.split(' - ')[1]}</a></div>}
                        </div>
                      )}
                    </div>
                  </div>
                :
                officemanager && item[officemanager.split('+')[0]] && item[officemanager.split('+')[0]].length > 0 &&
                  <div className={item[officemanager.split('+')[0]].length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{officemanager.split('+')[1]}</h3>
                    <div className={styles.employeeList}>

                      {item[officemanager.split('+')[0]].map((officemanager1: { Email: string; LastName: string; FirstName: string; JobTitle: string; EMail: string; },index: React.Key | null | undefined) => 
                      
                        <SPFxPeopleCard
                          email={officemanager1.EMail}
                          primaryText={`${officemanager1.LastName}, ${officemanager1.FirstName}`}
                          secondaryText={officemanager1.JobTitle}
                          serviceScope={webPartContext.serviceScope}
                          size={PersonaSize.size48}
                          key={`peopleCard-${officemanager1.EMail}`}
                          class={item[officemanager.split('+')[0]].length === 1 ? styles.employee100 : styles.employee50}
                        />

                      )}

                    </div>
                  </div>
                }

                {otherkeycontactstext && item[otherkeycontactstext.split('+')[0]] && item[otherkeycontactstext.split('+')[0]].split(';').length > 0 ?
                  <div className={item[otherkeycontactstext.split('+')[0]].split(';').length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{otherkeycontacts.split('+')[1]}</h3>
                    <div className={styles.employeeList}>
                      {item[otherkeycontactstext.split('+')[0]].split(';').map((lead: any,index: React.Key | null | undefined) => 
                        <div key={index} className={item[otherkeycontactstext.split('+')[0]].split(';').length === 1 ? styles.employee100 : styles.employee50}>
                          <div>{lead.split(' - ')[0]}</div>
                          {lead.split(' - ').length > 1 && <div><a href={`mailto:${lead.split(' - ')[1]}`}>{lead.split(' - ')[1]}</a></div>}
                        </div>
                      )}
                    </div>
                  </div>
                :
                otherkeycontacts && item[otherkeycontacts.split('+')[0]] && item[otherkeycontacts.split('+')[0]].length > 0 &&
                  <div className={item[otherkeycontacts.split('+')[0]].length === 1 ? styles.width50 : styles.width100} key={index}>
                    <h3>{otherkeycontacts.split('+')[1]}</h3>
                    <div className={styles.employeeList}>

                      {item[otherkeycontacts.split('+')[0]].map((otherkeycontact: { Email: string; LastName: string; FirstName: string; JobTitle: string; EMail: string; },index: React.Key | null | undefined) => 
                      
                        <SPFxPeopleCard
                          email={otherkeycontact.EMail}
                          primaryText={`${otherkeycontact.LastName}, ${otherkeycontact.FirstName}`}
                          secondaryText={otherkeycontact.JobTitle}
                          serviceScope={webPartContext.serviceScope}
                          size={PersonaSize.size48}
                          key={`peopleCard-${otherkeycontact.EMail}`}
                          class={item[otherkeycontacts.split('+')[0]].length === 1 ? styles.employee100 : styles.employee50}
                        />

                      )}

                    </div>
                  </div>
                }
                
              </div>
              
              {item[button1.split('+')[0]] && <a href={item[button1.split('+')[0]].split(';')[2]} className={styles.officeBtn} target="_blank" rel="noreferrer"><Icon iconName={item[button1.split('+')[0]].split(';')[0]} /> {item[button1.split('+')[0]].split(';')[1]}</a>}
              {item[button2.split('+')[0]] && <a href={item[button2.split('+')[0]].split(';')[2]} className={styles.officeBtn} target="_blank" rel="noreferrer"><Icon iconName={item[button2.split('+')[0]].split(';')[0]} /> {item[button2.split('+')[0]].split(';')[1]}</a>}
              {item[button3.split('+')[0]] && <a href={item[button3.split('+')[0]].split(';')[2]} className={styles.officeBtn} target="_blank" rel="noreferrer"><Icon iconName={item[button3.split('+')[0]].split(';')[0]} /> {item[button3.split('+')[0]].split(';')[1]}</a>}

            </div>
          )}
        </div>
      </section>
    );
  }

  public componentDidMount(): void {
    if(this.props.listId){
      this.getItems();
    }
  }

  public getItems(): void {

    this.setState({
      loading: true
    });

    const {
      listId,
      region,
      regionfield,
      country,
      countryflag,
      city,
      leads,
      hrlead,
      itlead,
      officemanager,
      otherkeycontacts,
      leadstext,
      hrleadtext,
      itleadtext,
      officemanagertext,
      otherkeycontactstext,
      button1,
      button2,
      button3
    } = this.props;

    const selectedFields = [];
    const selectedPeopleFields = [];
    const filter = region && `&$filter=${regionfield.split('+')[0]} eq '${region}'`
    const sorting = country && city ? `&$orderby=${country.split('+')[0]} asc, ${city.split('+')[0]} asc` : ''

    regionfield && selectedFields.push(regionfield.split('+')[0]);
    country && selectedFields.push(country.split('+')[0]);
    countryflag && selectedFields.push(countryflag.split('+')[0]);
    city && selectedFields.push(city.split('+')[0]);
    leads && selectedFields.push(`${leads.split('+')[0]}/LastName,${leads.split('+')[0]}/FirstName,${leads.split('+')[0]}/JobTitle,${leads.split('+')[0]}/EMail`); selectedPeopleFields.push(leads.split('+')[0]);
    hrlead && selectedFields.push(`${hrlead.split('+')[0]}/LastName,${hrlead.split('+')[0]}/FirstName,${hrlead.split('+')[0]}/JobTitle,${hrlead.split('+')[0]}/EMail`); selectedPeopleFields.push(hrlead.split('+')[0]);
    itlead && selectedFields.push(`${itlead.split('+')[0]}/LastName,${itlead.split('+')[0]}/FirstName,${itlead.split('+')[0]}/JobTitle,${itlead.split('+')[0]}/EMail`); selectedPeopleFields.push(itlead.split('+')[0]);
    officemanager && selectedFields.push(`${officemanager.split('+')[0]}/LastName,${officemanager.split('+')[0]}/FirstName,${officemanager.split('+')[0]}/JobTitle,${officemanager.split('+')[0]}/EMail`); selectedPeopleFields.push(officemanager.split('+')[0]);
    otherkeycontacts && selectedFields.push(`${otherkeycontacts.split('+')[0]}/LastName,${otherkeycontacts.split('+')[0]}/FirstName,${otherkeycontacts.split('+')[0]}/JobTitle,${otherkeycontacts.split('+')[0]}/EMail`); selectedPeopleFields.push(otherkeycontacts.split('+')[0]);
    leadstext && selectedFields.push(leadstext.split('+')[0]);
    hrleadtext && selectedFields.push(hrleadtext.split('+')[0]);
    itleadtext && selectedFields.push(itleadtext.split('+')[0]);
    officemanagertext && selectedFields.push(officemanagertext.split('+')[0]);
    otherkeycontactstext && selectedFields.push(otherkeycontactstext.split('+')[0]);
    button1 && selectedFields.push(button1.split('+')[0]);
    button2 && selectedFields.push(button2.split('+')[0]);
    button3 && selectedFields.push(button3.split('+')[0]);

    SharePointService.getListItems(listId, selectedFields, `&$top=500&$expand=${selectedPeopleFields.join(',')}${filter}${sorting}`).then(items => {

      this.setState({
        items: items.value,
        loading: false,
        error: 'null'
      });

    }).catch(error => {
        this.setState({
            error: 'Something went wrong!',
            loading: false
        });
    });
  }
}