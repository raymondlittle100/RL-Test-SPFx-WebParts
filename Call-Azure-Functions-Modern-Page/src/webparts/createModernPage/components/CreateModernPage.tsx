import * as React from 'react';
import styles from './CreateModernPage.module.scss';
import { ICreateModernPageProps } from './ICreateModernPageProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';

export default class CreateModernPage extends React.Component<ICreateModernPageProps, void> {


protected runFunction(): void {
  const requestHeaders: Headers = new Headers();
      requestHeaders.append("Content-type", "application/json");
      requestHeaders.append("Cache-Control", "no-cache");

      var siteUrl: string = this.props.siteUrl;
      let pageName: HTMLInputElement = ((HTMLInputElement)document.getElementById("txtPageName"));
      let pageText: HTMLInputElement = (HTMLInputElement)document.getElementById("txtPageText");
      var pageNameValue: string =  pageName.value;
      var pageTextValue: string = pageText.innerText;
debugger;
      console.log(`SiteUrl: '${siteUrl}', PageName: '${pageNameValue}.aspx', PageText: '${pageTextValue}'`);

      const postOptions: IHttpClientOptions = {
        headers: requestHeaders,
        body: `{ SiteUrl: '${siteUrl}', PageName: '${pageNameValue}.aspx', PageText: '${pageTextValue}' }`
      };

      let responseText: string = "";
      let resultMsg: HTMLElement = document.getElementById("responseContainer");

      this.props.httpClient.post(this.props.functionUrl, HttpClient.configurations.v1, postOptions).then((response: HttpClientResponse) => {
         response.json().then((responseJSON: JSON) => {
            responseText = JSON.stringify(responseJSON);
            if (response.ok) {
                resultMsg.style.color = "green";
            } else {
                resultMsg.style.color = "red";
            }

            resultMsg.innerText = responseText;
          })
          .catch ((response: any) => {
            let errMsg: string = `WARNING - error when calling URL ${this.props.functionUrl}. Error = ${response.message}`;
            resultMsg.style.color = "red";
            console.log(errMsg);
            resultMsg.innerText = errMsg;
          });
});
}

  public render(): React.ReactElement<ICreateModernPageProps> {
    return (
      <div className={styles.createModernPage}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">Create a new page in this site</span>
              <div className="${styles.formRow}">
                <span className="ms-font-l ms-fontColor-white ${styles.formLabel}">Page name:</span>
                <input type="text" ref="txtPageName" id="txtPageName"></input>.aspx
              </div>
              <div className="${styles.formRow}">
                <span className="ms-font-l ms-fontColor-white ${styles.formLabel}">Page content:</span>
                <input type="text" ref={(input) => this.input = input}></input>
              </div>
              <div className="${styles.buttonRow}"></div>
              <button id="btnCallFunction" onClick={() => this.runFunction() } >Call Function</button>
              <div id="responseContainer" className="${styles.result}"></div>              
            </div>
          </div>
        </div>
      </div>
    );
  }
}
