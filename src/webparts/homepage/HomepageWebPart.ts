import { Version } from '@microsoft/sp-core-library';
import $ from 'jquery';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HomepageWebPart.module.scss';
import * as strings from 'HomepageWebPartStrings';
import 'bootstrap/dist/js/bootstrap.bundle.min';
import { Navigation } from 'spfx-navigation';

import 'bootstrap/dist/css/bootstrap.css';
import { SPHttpClientResponse } from '@microsoft/sp-http';
import { SPHttpClient } from '@microsoft/sp-http';
import * as _ from 'lodash';
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/jquery/3.4.1/jquery.min.js");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.6.1/gsap.min.js");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/1.20.2/TweenMax.min.js");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.6.1/CSSRulePlugin.min.js");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.6.1/ScrollTrigger.min.js");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/Swiper/5.1.0/js/swiper.min.js");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.6.1/ScrollTrigger.min.js");


// require('./../../../lib/common/css/bootstrap/mi');
// require('./../../../common/css/basic.css');
require('./../../../src/common/css/media.css');
require('./../../../src/common/css/basic.css');
require('./../../../src/common/css/global.css');
require('./../../../src/common/css/common.css');
// require('./../../../src/common/css/qlf5ifj.css');
require('./../../../src/common/js/custom.js');
require('./../../../src/common/js/animation.js');




export interface IHomepageWebPartProps {
    description: string;
}

export default class HomepageWebPart extends BaseClientSideWebPart<IHomepageWebPartProps> {
    [x: string]: any;

    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';
    


    protected onInit(): Promise<void> {
        // this._environmentMessage = this._getEnvironmentMessage();

        return super.onInit();
    }

    public render(): void {


        this.domElement.innerHTML = ` 
        <main>
    <div class="main-container w100">
        <section class="banner-section w100">
            <div class="photo w100" id="navImage"> 
 
            </div>
        </section>

        <section class="cta-mg-section w100">
            <div class="inner-ctamg-section w100 cnt-80 flex-basic">
                <div class="cta-mg-repeated">
                    <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta1.png")}')">
                        Divers
                    </a>
                </div>

                <div class="cta-mg-repeated">
                    <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta2.png")}')">
                        DOCUMENTATION
                    </a>
                </div>

                <div class="cta-mg-repeated">
                    <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta3.png")}')">
                        CRISE
                    </a>
                </div>

                <div class="cta-mg-repeated">
                    <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta4.png")}')">
                        PNT
                    </a>
                </div>
            </div>
        </section>

        <section class="mg-text-section w100 pda-75">
            <div class="inner-mg-text w100 cnt-85 pda-50">
                <div class="mg-text-bloc w100 cnt-95 flex-basic">
                    <div class="mg-text-repeated">
                        <!-- for mobile only -->
                        <div class="cta-mg-repeated">
                            <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta1.png")}')">
                                Divers
                            </a>
                        </div>
                        <!-- end for mobile only -->
                  
                        <div class="mg-cta-repeated w100" id="list1">

                        </div>

                       
                    </div>
                    <div class="mg-text-repeated">
                        <!-- for mobile only -->
                        <div class="cta-mg-repeated">
                            <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta2.png")}')">
                                Divers
                            </a>
                        </div>
                        <!-- end for mobile only -->


                        <div class="mg-cta-repeated w100" id="list2">

                 
                        
        
                        </div>

                    </div>
                    <div class="mg-text-repeated">
                        <!-- for mobile only -->
                        <div class="cta-mg-repeated">
                            <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta3.png")}')">
                                Divers
                            </a>
                        </div>
                        <!-- end for mobile only -->
             

                        <div class="mg-cta-repeated w100" id="list3">

                        </div>


                    </div>   
                    <div class="mg-text-repeated">
                        <!-- for mobile only -->
                        <div class="cta-mg-repeated">
                            <a href="javascript:void(0)" style="background-image: url('${require<string>("./../../common/images/bg-cta4.png")}')">
                                Divers
                            </a>
                        </div>
                        <!-- end for mobile only -->
                   

                        <div class="mg-cta-repeated w100" id="list4">

                        </div>
                        
                        
                    </div>
                </div>
            </div>
        </section>
    </div>

    <footer class="w100">
        <div class="footer-top w100 cnt-75">
            © 2022 MyAircalin
        </div>



        <div class="footer-bottom w100">
            <img src= "${require<string>('./../../common/images/img-footer-bottom.png')}" class="img-responsive" alt="">
        </div>
    </footer>
</main>`;

        this.eventTriggers();
        this._renderNavImage();
         this._getBuildingsList();

        // this._gethomepageLinks();
        //this._renderHomepageLinks();
        // this._renderHomepageLinks2();
        // this._renderHomepageLinks3();
        // this._renderHomepageLinks4();
    }

    private eventTriggers() {

        $(".info-emploi-title").on("click", () => {
            Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx?folder=2`, true);
        });
    }


    //API to get navImage
    private async _getNavImage(): Promise<any> {
        const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('PageDaccueilPhoto')/Items", SPHttpClient.configurations.v1);
        return await response.json();
    }

    private _renderNavImage(): void {
        const listContainerImage: Element = this.domElement.querySelector('#navImage');
        this._getNavImage().then(async (response) => {
            console.log(response.value);

            await Promise.all(response.value.map(async (result) => {
                let html: string = ''

                const item = {
                    Title: result.Title,
                    Image: result.Image
                };

                const imageJson = ((JSON.parse(item.Image)).serverRelativeUrl);


                html += `<img src="https://ncaircalin.sharepoint.com/${imageJson}" class="img-responsive" alt="" />`


                listContainerImage.innerHTML += html;
            }
            ))
        });
    }

    //API to get homepageLinks
    private async _gethomepageLinks(): Promise<any> {

        const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('HomepageLinks')/Items?$select=Title,url,Order,permission,linksType", SPHttpClient.configurations.v1);

        response.json().then(async (response) => {
            this.arrayLinks = response.value;
            console.log("ARRAYLINKS", this.arrayLinks);
            return this.arrayLinks;
        });
    }

    private _getBuildingsList() {

  var arrayLinks: any[];

        let html1: string = '';
        let html2: string = '';
        let html3: string = '';
        let html4: string = '';

        return new Promise(async (resolve, reject) => {
            try {
                this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('HomepageLinks')/items`, SPHttpClient.configurations.v1)
                    .then(response => {
                        return response.json()
                            .then((items: any): void => {
                                arrayLinks = items.value;

                                console.log("ARRAYLINKS", arrayLinks);

                                // if ((item.Order === 1) || (item.Order === 2 )|| ( item.Order === 3 ) || (item.Order === 4) || (item.Order === 5 ) ) {
                               

                                arrayLinks.forEach((item: any) => {

                                    console.log("URL", item.url);

                                    if ((item.order0 == "1") || (item.order0 == "2" )|| ( item.order0 == "3" ) || (item.order0 == "4") || (item.order0 == "5" ) ) {

                                        console.log("ORDER 1-5", item.Title);
                                        
                                        html1 += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                ${item.Title}
                                            </div>
                                        </div>`;

                                    

                                    }

                                    // else if ((item.Order === 6) || (item.Order === 7) || (item.Order === 8) || (item.Order === 9)) {

                                    else if ((item.order0 == "6") || (item.order0 == "7") || (item.order0 == "8") || (item.order0 == "9")) {
                                        html2 += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                ${item.Title}
                                            </div>
                                        </div>`;

                                    }
                                   // else if ((item.Order === 10) || (item.Order === 11)) {

                                     else if ((item.order0 == "10") || (item.order0 == "11")) {
                                        html3 += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                ${item.Title}
                                            </div>
                                        </div>`;

                                    }

                                    else {


                                        html4 += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                ${item.Title}
                                            </div>
                                        </div>`;

                                    }

                                });

                                const listContainer1: Element = this.domElement.querySelector('#list1');
                                listContainer1.innerHTML += html1;
    
                                const listContainer2: Element = this.domElement.querySelector('#list2');
                                listContainer2.innerHTML += html2;
    
                                const listContainer3: Element = this.domElement.querySelector('#list3');
                                listContainer3.innerHTML += html3;
    
                                const listContainer4: Element = this.domElement.querySelector('#list4');
                                listContainer4.innerHTML += html4;
    
                            });
                    });

            }
            catch (error) {
                console.log(error);
                reject(error);
            }
        });

    }



    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }

        this._isDarkTheme = !!currentTheme.isInverted;
        const {
            semanticColors
        } = currentTheme;
        this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
        this.domElement.style.setProperty('--link', semanticColors.link);
        this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
