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
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item, Web } from "@pnp/sp/presets/all";


import styles from './HomepageWebPart.module.scss';
import * as strings from 'HomepageWebPartStrings';
import 'bootstrap/dist/js/bootstrap.bundle.min';
// import { Navigation } from 'spfx-navigation';

import 'bootstrap/dist/css/bootstrap.css';

import Swiper, { Navigation, Pagination, Grid, Autoplay, EffectFade } from 'swiper';

Swiper.use([Navigation, Pagination, Grid, Autoplay, EffectFade]);





import * as _ from 'lodash';
// import { sp } from '@pnp/sp';




<<<<<<< HEAD

=======
>>>>>>> fdf4fd44f05486068e109e7465cdfcc07e607039



export interface IHomepageWebPartProps {
    description: string;
}

export default class HomepageWebPart extends BaseClientSideWebPart<IHomepageWebPartProps> {
    [x: string]: any;

    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';



    protected onInit(): Promise<void> {
        // this._environmentMessage = this._getEnvironmentMessage();
        sp.setup({
            spfxContext: this.context
        });

        return super.onInit();
    }

    public render(): void {



        this.domElement.innerHTML = ` 
        <main>
    <div class="main-container w100">

    <section class="banner-section w100">
    <div class="swiper">

    <div class="swiper-wrapper" id="swiper_image">


    </div>


    <div class="swiper-button-prev"></div>
    <div class="swiper-button-next"></div>

</div>  
</section>

      

        <section class="cta-mg-section w100">
            <div class="inner-ctamg-section w100 cnt-80 flex-basic" style="z-index: 121;" id="catBtn">
                
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
</main>

`;
        // require('./../../../lib/common/css/bootstrap/mi');
        // require('./../../../common/css/basic.css');
        require('./../../../src/common/css/media.css');
        require('./../../../src/common/css/basic.css');
        require('./../../../src/common/css/global.css');
        require('./../../../src/common/css/common.css');
        // require('./../../../src/common/css/qlf5ifj.css');
        require('./../../../src/common/js/custom.js');
        setTimeout(() => {
            require('./../../../src/common/js/animation.js');
<<<<<<< HEAD
        }, 2000);

=======
        }, 2000)
>>>>>>> fdf4fd44f05486068e109e7465cdfcc07e607039

        SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/jquery/3.4.1/jquery.min.js");
        SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.6.1/gsap.min.js");
        SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/1.20.2/TweenMax.min.js");
        SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.6.1/CSSRulePlugin.min.js");
        SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.6.1/ScrollTrigger.min.js");
        SPComponentLoader.loadCss('https://unpkg.com/swiper@7/swiper-bundle.min.css');


        this._renderNavImage();
        this._renderCatBtn();
        this._getBuildingsList();

    }


    //API to get navImage
    private async _getNavImage(): Promise<any> {
        //  const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('PageDaccueilPhoto')/Items", SPHttpClient.configurations.v1);
        const response: any[] = (await sp.web.lists.getByTitle("PageDaccueilPhoto").items.orderBy("Order", false).get());
        // console.log("order", response.sort);
        return response;
        // return await response.json();
    }


    private _renderNavImage(): void {

        const listContainerImage: Element = this.domElement.querySelector('.swiper-wrapper');

        let swiper_html: string = '';

        this._getNavImage().then(async (response) => {
            // console.log("IMAGE", response);

            response.forEach((item) => {
                const imageJson = ((JSON.parse(item.Image)).serverRelativeUrl);
                console.log("JSONIMAGE", imageJson);
                let html = ` <div class="swiper-slide">
<<<<<<< HEAD
                    <div style="background-image: url('https://ncaircalin.sharepoint.com${imageJson}'); background-position: ${item.PositionVerticale}% ${item.PositionHorizontale}%" class="banner">
=======
                    <div style="background-image: url('https://ncaircalin.sharepoint.com/${imageJson}'); background-position: ${item.PositionVerticale}% ${item.PositionHorizontale}%" class="banner">
>>>>>>> fdf4fd44f05486068e109e7465cdfcc07e607039
                    </div>
                </div>`;
                swiper_html += html;
            });
            listContainerImage.innerHTML = swiper_html;
        })
            .then(() => {
                this._swipe();
            });
    }


    private _swipe() {
        const swipercol = new Swiper(".swiper", {
            slidesPerView: 1,
            effect: 'fade',
            fadeEffect: {
                crossFade: true
            },
            loop: true,
            navigation: {
                nextEl: ".swiper-button-next",
                prevEl: ".swiper-button-prev",
            },

            autoplay: {
                delay: 5000,
                disableOnInteraction: false,
                pauseOnMouseEnter: true,
            }
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

                const response: any[] = await sp.web.lists.getByTitle("HomepageLinks").items();
                // console.log("RESPONSE", response);
                arrayLinks = response;
                // if ((item.Order === 1) || (item.Order === 2 )|| ( item.Order === 3 ) || (item.Order === 4) || (item.Order === 5 ) ) {
                arrayLinks.forEach((item: any) => {
                    // console.log("URL", item.url);
                    // if ((item.order0 == "1") || (item.order0 == "2") || (item.order0 == "3") || (item.order0 == "4") || (item.order0 == "5")) {
                    if ((item.linksType == "Divers")) {
                        // console.log("ORDER 1-5" , item.title);
                        html1 += `<a href="${item.url}" class="w100 flex-basic flex-justify-center flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                > ${item.Title}
                                            </div>
                                        </div>`;

                    }
                    // else if ((item.Order === 6) || (item.Order === 7) || (item.Order === 8) || (item.Order === 9)) {
                    // else if ((item.order0 == "6") || (item.order0 == "7") || (item.order0 == "8") || (item.order0 == "9")) {
                    else if ((item.linksType == "Documentation")) {
                        html2 += `<a href="${item.url}" class="w100 flex-basic flex-justify-center flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                > ${item.Title}
                                            </div>
                                        </div>`;

                    }
                    // else if ((item.Order === 10) || (item.Order === 11)) {
                    // else if ((item.order0 == "10") || (item.order0 == "11")) {
                    else if ((item.linksType == "Crise")) {
                        html3 += `<a href="${item.url}" class="w100 flex-basic flex-justify-center flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                > ${item.Title}
                                            </div>
                                        </div>`;
                    }
                    else {
                        if ((item.linksType == "PNT")) {
                            html4 += `<a href="${item.url}" class="w100 id="PNT" flex-basic flex-justify-center flex-align-center">
                                        <div class="info-emploi-text w85">
                                            <div class="info-emploi-title">
                                                > ${item.Title}
                                            </div>
                                        </div>`;
                        }
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

                // });
                // });

            }
            catch (error) {
                console.log(error);
                reject(error);
            }
        });

    }


    // private async checkPermission() {
    //     const groupTitle = [];
    //     let groups: any = await sp.web.currentUser.groups();

    //     await Promise.all(groups.map(async (perm) => {

    //         groupTitle.push(perm.Title);

    //     }));

    //     // if (groupTitle.includes("myGed Visitors")) {
    //     if (groupTitle.includes("myGED Visitors")) {

    //         $("#PNT").css("display", "none");

    //         $("#PNT ").prop('disabled', true);

    //     }
    //     else {

    //         // $("#update_details_doc, #edit_cancel_doc, #access, #notifications, #audit").css("display", "block");
    //     }


    // }

    private async _renderCatBtn() {
        let web = Web(this.context.pageContext.web.absoluteUrl);
        const items = await web.lists.getByTitle("HomepageCatergoryLinks").items.get();
        let htmlcatBtn = '';
        const catBtn: Element = this.domElement.querySelector('#catBtn');
        items.forEach((element) => {
            element = {
                Title: element.Title,
                url: element.url,
                bgImage: element.bgImage
            };
            const imageJson2 = ((JSON.parse(element.bgImage)).serverRelativeUrl);
<<<<<<< HEAD
            if (imageJson2 != null) {
=======
            if(imageJson2 != null){
>>>>>>> fdf4fd44f05486068e109e7465cdfcc07e607039
                htmlcatBtn += `
                <div class="cta-mg-repeated">
                    <a href="${element.url}" style="background-image: url('https://ncaircalin.sharepoint.com/${imageJson2}')">
                        ${element.Title}
                    </a>
                </div>`;
            }
<<<<<<< HEAD
            else if (imageJson2 == null) {
=======
            else if (imageJson2 == null){
>>>>>>> fdf4fd44f05486068e109e7465cdfcc07e607039
                htmlcatBtn += `
                <div class="cta-mg-repeated">
                    <a href="${element.url}" style="background-image: url('${require<string>('./../../common/images/bg-cta4.png')}')">
                        ${element.Title}
                    </a>
                </div>`;
            }
            console.log("url", element.url);
        });
        catBtn.innerHTML += htmlcatBtn;
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

    // protected get dataVersion(): Version {
    //         return Version.parse('1.0');
    //     }

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
