import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CarouselBenefitsWebPart.module.scss';
import * as strings from 'CarouselBenefitsWebPartStrings';

import $ from "jquery";
import * as bootstrap from 'bootstrap';

require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/bootstrap/js/dist/carousel.js');

import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { ItemActionStat } from '@microsoft/microsoft-graph-types';

export interface IGetSpListItemsWebPartProps {
  URLTenant: string;
  listName: string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Beneficio: string;
  Nombre_x0020_de_x0020_la_x0020_e: string;
  Fecha_x0020_de_x0020_vencimiento: string;
  Requisitos: string;
  Sucursales: string;
  Tipo_x0020_de_x0020_beneficio: string;
  Pagina_x0020_web: string;
  Facebook: string;
  Instagram: string;
  Logo: any;
  Contacto: string;
  Fecha_x0020_de_x0020_vencimiento0: string;
}
  
export default class GetSpListItemsWebPart extends BaseClientSideWebPart<IGetSpListItemsWebPartProps> {
  
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.properties.URLTenant + `/_api/web/lists/GetByTitle('${this.properties.listName}')/Items`,SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
        return response.json();
        });
    }

  private _renderListAsync(): void {
    
      if (Environment.type == EnvironmentType.SharePoint || 
               Environment.type == EnvironmentType.ClassicSharePoint) {
       this._getListData()
         .then((response) => {
           this._renderList(response.value);
         });
     }
   }


    private _renderList(items: ISPList[]): void {

      console.log(items);

      let numberRows = (items.length / 5);
      let counter: number = 0;


      for(let _i = 0; _i <= numberRows; _i++){
 
        let content: string = '';
        // let indicator: string = '';

        let rows: string = '';

        rows += `
        <div class="carousel-item ${_i == 0 ? "active" : ""}">
          <div class="row" id="Value${_i}" style="box-sizing: border-box; width: auto;" >
          </div>
        </div>
        `

        for(let _x = counter; _x <= (counter + 4); _x++){

          if(items[_x] != null || items[_x] != undefined){
            
            // indicator += `
            //   <li data-target="#carouselBenefits" data-slide-to="${_x}" class="${_x == 0 && _i == 0 ? 'active' : ""}"></li>
            // `   
            content += `
              <div class="col"> 
                <a href="https://grupoporres.sharepoint.com/sites/Intranet/Lists/Beneficios%20corporativos/AllItems.aspx?FilterField1=Nombre%5Fx0020%5Fde%5Fx0020%5Fla%5Fx0020%5Fe&FilterValue1=${items[_x].Nombre_x0020_de_x0020_la_x0020_e}&FilterType1=Text&FilterOp1=In" target="_blank"> 
                  <img style="margin: 10px 20px;" width="100" src="${items[_x].Logo.Url}" alt="${items[_x].Nombre_x0020_de_x0020_la_x0020_e}">
                </a>
              </div>    
            `                
          }
        
        }

        counter += 5;

        const listRows: Element = this.domElement.querySelector('.carousel-inner');
        listRows.insertAdjacentHTML('beforeend', rows);

        // const listIndicator: Element = this.domElement.querySelector('.carousel-indicators');
        // listIndicator.insertAdjacentHTML('beforeend', indicator);

        const listContent: Element = this.domElement.querySelector(`#Value${_i}`);
        listContent.insertAdjacentHTML('beforeend', content);
      }



    }
      
    // <ol class="carousel-indicators">
    // </ol>

  public render(): void {
    this.domElement.innerHTML = `

    <div id="carouselBenefits" class="carousel slide" data-ride="carousel" style="width: 900px">
      <div class="carousel-inner">

      </div>
      <a class="carousel-control-prev" href="#carouselBenefits" role="button" data-slide="prev">
        <span class="carousel-control-prev-icon" aria-hidden="true"></span>
        <span class="sr-only">Previous</span>
      </a>
      <a class="carousel-control-next" href="#carouselBenefits" role="button" data-slide="next">
        <span class="carousel-control-next-icon" aria-hidden="true"></span>
        <span class="sr-only">Next</span>
      </a>
    </div>`;


      this._renderListAsync();


  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('URLTenant', {
                  label: 'URL Tenant:'
                }),
                PropertyPaneTextField('listName', {
                  label: 'Nombre de la lista:'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}