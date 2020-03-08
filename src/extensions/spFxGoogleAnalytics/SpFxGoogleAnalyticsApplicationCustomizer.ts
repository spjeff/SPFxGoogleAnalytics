import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'SpFxGoogleAnalyticsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SpFxGoogleAnalyticsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFxGoogleAnalyticsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  trackingID: string;
  MissingID: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpFxGoogleAnalyticsApplicationCustomizer extends BaseApplicationCustomizer<ISpFxGoogleAnalyticsApplicationCustomizerProperties> {

    // from https://www.sharepointvitals.com/blog/google-analytics-for-sharepoint-ultimate-guide/

    private currentPage = "";


  private isInitialLoad = true;

  private getFreshCurrentPage(): string {

    return window.location.pathname + window.location.search;

  }

  private updateCurrentPage(): void {

    this.currentPage = this.getFreshCurrentPage();

  }

  private navigatedEvent(): void {

    let trackingID: string = this.properties.trackingID;

    if (!trackingID) {

      Log.info(LOG_SOURCE, `${strings.MissingID}`);

    } else {

      const navigatedPage = this.getFreshCurrentPage();

      if (this.isInitialLoad) {

        this.realInitialNavigatedEvent(trackingID);

        this.updateCurrentPage();

        this.isInitialLoad = false;

      }

      else if (!this.isInitialLoad && (navigatedPage !== this.currentPage)) {

        this.realNavigatedEvent(trackingID);

        this.updateCurrentPage();

      }

    }

  }

  private realInitialNavigatedEvent(trackingID: string): void {

    console.log("Tracking full page load...");

    var gtagScript = document.createElement("script");

    gtagScript.type = "text/javascript";

    gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${trackingID}`;

    gtagScript.async = true;

    document.head.appendChild(gtagScript);

    eval(`

        window.dataLayer = window.dataLayer || [];

        function gtag(){dataLayer.push(arguments);}

        gtag('js', new Date());

        gtag('config',  '${trackingID}');

      `);

  }

  private realNavigatedEvent(trackingID: string): void {

    console.log("Tracking partial page load...");

    eval(`

      if(ga) {

        ga('create', '${trackingID}', 'auto');

        ga('set', 'page', '${this.getFreshCurrentPage()}');

        ga('send', 'pageview');

      }

      `);

  }

  @override

  public onInit(): Promise<any> {

    this.context.application.navigatedEvent.add(this, this.navigatedEvent);

    return Promise.resolve();

  }
}
