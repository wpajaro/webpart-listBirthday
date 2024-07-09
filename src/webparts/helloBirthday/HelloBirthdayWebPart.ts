
//import * as React from 'react';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './HelloBirthdayWebPart.module.scss';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

export interface IHelloBirthdayWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  NombresApellidos: string;
}

export default class HelloBirthdayWebPart extends BaseClientSideWebPart<IHelloBirthdayWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloBirthday}">
      <div class="${styles.welcome}">
        <h2> Cumpleaños Oficina Asesora de Informática </h2>
      </div>
      <div>
        <h4  class="${styles.tituloParrafo}"> ¡Celebramos un cumpleaños hoy! </h4>
        <p class="${styles.parrafo}"> Hoy es un día especial para uno de nuestros compañeros. 
        ¡Deseémosle un feliz cumpleaños y que tenga un día maravilloso 
        lleno de alegría y buenos momentos! </p>
      </div>
      <div id="spListContainer"/>
    </section>`;

  this._renderListAsync();
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      })
      .catch((error) => {
        console.error('Error al obtener los datos de la lista', error);
        throw error;
      });
  }

  //Obtiene el acceso a todas las listas creadas en SharePoint  
  private _getListData(): Promise<ISPLists> {
    const listName = "Contratista";
    
    const date: Date = new Date()
    const month: string = `${date.getMonth()+1}`
    const day: string = `${date.getDate()}`

    const fechaString: string = day + '-' + month 
    const nuevaFecha: string = fechaString  
    
    
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=fechaString eq '${nuevaFecha}'`;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .catch((error) => {
        console.error('Error al obtener los datos de la lista', error);
        throw error;
      });
  }

  private _renderList(items: ISPList[]): void {
    
    let html: string = '';
    items.forEach((item: ISPList) => {
        html += `
          <ul class="${styles.list}">
           <li class="${styles.listItem}">
              <span class="ms-font-1">${item.NombresApellidos}</span>
            </li>
           </ul>`;   
    });

    const listContainer = this.domElement.querySelector('#spListContainer');
    if(listContainer){
      listContainer.innerHTML = html;
    }
  }


}

