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
import { SPHttpClient } from '@pnp/sp';
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
require('./../../../src/common/css/qlf5ifj.css');
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
        this._environmentMessage = this._getEnvironmentMessage();

        return super.onInit();
    }

    public render(): void {
        this.domElement.innerHTML = ` <main>
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
                        <div id="homepageLinksDivers1">
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
                        <div id="homepageLinksDivers2">
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
                        <div id="homepageLinksDivers3">
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
                        <div id="homepageLinksDivers4">
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
        this._renderHomepageLinks();
        this._renderHomepageLinks2();
        this._renderHomepageLinks3();
        this._renderHomepageLinks4();
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
            var navImage = [];
            await Promise.all(response.value.map(async (result: { Title: any; Image: any; }) => {
                let html: string = ''

                const item = {
                    Title: result.Title,
                    Image: result.Image
                };

                await navImage.reduce(async (memo, item) => {
                    await memo;
                    const imageJson = ((JSON.parse(item.Image)).serverRelativeUrl)

                    html += `<img src="https://ncaircalin.sharepoint.com/${imageJson}" class="img-responsive" alt="" />`

                    listContainerImage.innerHTML += html;
                })
            }
            ))
        });
    }

    //API to get homepageLinks
    private async _gethomepageLinks(): Promise<any> {
        const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('HomepageLinks')/Items?$select=Title,url,order,permission,linksType&$filter=order eq 1,2,3,4,5", SPHttpClient.configurations.v1);
        return await response.json();
    }

    private _renderHomepageLinks() {
        const listContainerHomepageLinks: Element = this.domElement.querySelector('#homepageLinks');
        this._gethomepageLinks().then(async (response) => {
            console.log(response.value);
            await Promise.all(response.value.map(async (result: { Title: any; url: any; order: any; permission: any; linksType: any; }) => {
                let homepageLinkshtml: string = '<div class="mg-cta-repeated w100">'

                const item = {
                    Title: result.Title,
                    url: result.url,
                    order: result.order,
                    permission: result.permission,
                    linksType: result.linksType
                };

                homepageLinkshtml += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                    <div class="info-emploi-text w85">
                        <div class="info-emploi-title">
                            ${item.Title}
                        </div>
                    </div>

                    <div class="info-emplo-cta w10">
                        <div class="cta-arrow blue">
                            <span class="btn">
                                <span class="arrow"></span>
                            </span>
                        </div>
                    </div>
                </a>`
                homepageLinkshtml += `</div>`
                listContainerHomepageLinks.innerHTML += homepageLinkshtml;
            }))
        });
    }


    //API to get homepageLinks2
    private async _gethomepageLinks2(): Promise<any> {
        const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('HomepageLinks')/Items?$select=Title,url,order,permission,linksType&$filter=order eq 6,7,8,9", SPHttpClient.configurations.v1);
        return await response.json();
    }

    private _renderHomepageLinks2() {
        const listContainerHomepageLinks: Element = this.domElement.querySelector('#homepageLinksDiver2');
        this._gethomepageLinks2().then(async (response) => {
            console.log(response.value);
            await Promise.all(response.value.map(async (result: { Title: any; url: any; order: any; permission: any; linksType: any; }) => {
                let homepageLinkshtml2: string = '<div class="mg-cta-repeated w100">'

                const item = {
                    Title: result.Title,
                    url: result.url,
                    order: result.order,
                    permission: result.permission,
                    linksType: result.linksType
                };

                homepageLinkshtml2 += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                    <div class="info-emploi-text w85">
                        <div class="info-emploi-title">
                            ${item.Title}
                        </div>
                    </div>

                    <div class="info-emplo-cta w10">
                        <div class="cta-arrow blue">
                            <span class="btn">
                                <span class="arrow"></span>
                            </span>
                        </div>
                    </div>
                </a>`
                homepageLinkshtml2 += `</div>`
                listContainerHomepageLinks.innerHTML += homepageLinkshtml2;
            }))
        });
    }

    //API to get homepageLinks3
    private async _gethomepageLinks3(): Promise<any> {
        const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('HomepageLinks')/Items?$select=Title,url,order,permission,linksType&$filter=order eq 10,11", SPHttpClient.configurations.v1);
        return await response.json();
    }

    private _renderHomepageLinks3() {
        const listContainerHomepageLinks: Element = this.domElement.querySelector('#homepageLinksDiver3');
        this._gethomepageLinks3().then(async (response) => {
            console.log(response.value);
            await Promise.all(response.value.map(async (result: { Title: any; url: any; order: any; permission: any; linksType: any; }) => {
                let homepageLinkshtml3: string = '<div class="mg-cta-repeated w100">'

                const item = {
                    Title: result.Title,
                    url: result.url,
                    order: result.order,
                    permission: result.permission,
                    linksType: result.linksType
                };

                homepageLinkshtml3 += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                    <div class="info-emploi-text w85">
                        <div class="info-emploi-title">
                            ${item.Title}
                        </div>
                    </div>

                    <div class="info-emplo-cta w10">
                        <div class="cta-arrow blue">
                            <span class="btn">
                                <span class="arrow"></span>
                            </span>
                        </div>
                    </div>
                </a>`
                homepageLinkshtml3 += `</div>`
                listContainerHomepageLinks.innerHTML += homepageLinkshtml3;
            }))
        });
    }

    //API to get homepageLinks4
    private async _gethomepageLinks4(): Promise<any> {
        const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('HomepageLinks')/Items?$select=Title,url,order,permission,linksType&$filter=order eq 12", SPHttpClient.configurations.v1);
        return await response.json();
    }

    private _renderHomepageLinks4() {
        const listContainerHomepageLinks: Element = this.domElement.querySelector('#homepageLinksDiver3');
        this._gethomepageLinks4().then(async (response) => {
            console.log(response.value);
            await Promise.all(response.value.map(async (result: { Title: any; url: any; order: any; permission: any; linksType: any; }) => {
                let homepageLinkshtml4: string = '<div class="mg-cta-repeated w100">'

                const item = {
                    Title: result.Title,
                    url: result.url,
                    order: result.order,
                    permission: result.permission,
                    linksType: result.linksType
                };

                homepageLinkshtml4 += `<a href="${item.url}" class="w100 flex-basic flex-justify-between flex-align-center">
                    <div class="info-emploi-text w85">
                        <div class="info-emploi-title">
                            ${item.Title}
                        </div>
                    </div>

                    <div class="info-emplo-cta w10">
                        <div class="cta-arrow blue">
                            <span class="btn">
                                <span class="arrow"></span>
                            </span>
                        </div>
                    </div>
                </a>`
                homepageLinkshtml4 += `</div>`
                listContainerHomepageLinks.innerHTML += homepageLinkshtml4;
            }))
        });
    }




    private _getEnvironmentMessage(): string {
        if (!!this.context.sdks.microsoftTeams) { // running in Teams
            return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
        }

        return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
