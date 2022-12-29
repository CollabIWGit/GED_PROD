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
            <div class="photo w100">
            
                <img src="${require<string>('./../../common/images/img-banner-myGed.jpg')}" class="img-responsive" alt="" />
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

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Plaquette Qualité / Quality Overview
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        SGS / SMS
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Q-Pulse Training
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Compte-Rendus d'événements / Reporting forms
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Q-Pulse Reporting
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
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
                        
                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Consultation du Manex / Ops Manual
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title" id="manorg">
                                        Consultation du MANORG
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Consultation du GOM / Ground Ops Manual
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        AIR@V (réseau interne uniquement)
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
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
                        
                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Manuel ALEAS / Disruption Manual
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
                        </div>

                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Crise / Crisis
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
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
                        
                        <div class="mg-cta-repeated w100">
                            <a href="javascript:void(0)" class="w100 flex-basic flex-justify-between flex-align-center">
                                <div class="info-emploi-text w85">
                                    <div class="info-emploi-title">
                                        Informations ASV / Flight Strategy
                                    </div>
                                </div>

                                <div class="info-emplo-cta w10">
                                    <div class="cta-arrow blue">
                                        <span class="btn">
                                            <span class="arrow"></span>
                                        </span>
                                    </div>
                                </div>
                            </a>
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
  }

  private eventTriggers() {

    $("#manorg").on("click", () => {
      Navigation.navigate(`${this.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx?folder=1312`, true);
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
