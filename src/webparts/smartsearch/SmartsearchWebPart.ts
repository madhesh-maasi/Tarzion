import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./SmartsearchWebPart.module.scss";
import * as strings from "SmartsearchWebPartStrings";
import { sp } from "@pnp/sp/presets/all";
import { setup as pnpSetup } from "@pnp/common";
import "@pnp/sp/search";
import {
  ISearchQuery,
  SearchResults,
  SearchQueryBuilder,
} from "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "../../ExternalRef/css/style.css";
export interface ISmartsearchWebPartProps {
  description: string;
}

export default class SmartsearchWebPart extends BaseClientSideWebPart<ISmartsearchWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      // other init code may be present

      pnpSetup({
        spfxContext: this.context,
      });
    });
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="section">
      <div class="left">
      <ul id="titleList"><ul>
      </div>
      <div class="right"><ul></ul></div>
      </div>`;
    fetchDatas();
    // getSearchValue();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
async function fetchDatas() {
  let list = await sp.web.lists
    .getByTitle("Categories")
    .items.orderBy("Display_x0020_Index", true)
    .get();

  // TODO Bind Title

  let setTitle = "";
  list.forEach((item) => {
    item.Text_x0020_Black == true
      ? (setTitle += `<div class="title-container">
    <div class="arrow" style="border-top: 10px solid ${item.Colour}"></div>
    <li class="listTitle" style="background:${item.Colour};color:#000">${item.Title}</li></div>`)
      : (setTitle += `<div class="title-container">
    <div class="arrow" style="border-top: 10px solid ${item.Colour}"></div>
    <li class="listTitle" style="background:${item.Colour}; color:#fff">${item.Title}</li></div>`);
  });
  let ul = document.getElementById("titleList");
  ul.innerHTML = setTitle;

  // TODO Click Event

  let LiTitle = document.querySelectorAll(".listTitle");
  LiTitle.forEach((li) => {
    li.addEventListener("click", (e) => {
      console.log(e.currentTarget["textContent"]);
      console.log(e.target["textContent"]);
      document.querySelector(".right")["style"][
        "background"
      ] = `${e.target["style"]["background"]}`;
      let wantedTitle = e.currentTarget["textContent"];
      let activeClass = document.querySelectorAll("#titleList li");
      activeClass.forEach((activeC) => {
        activeC["classList"].remove("active");
      });
      e.target["classList"].toggle("active");

      // TODO Search

      sp.search({
        Querytext: `${e.currentTarget["textContent"]} path:https://chandrudemo.sharepoint.com/sites/Torsion/`,
        RowLimit: 100,
        // RefinementFilters: ["IsDocument:true"],
        TrimDuplicates: true,
        SelectProperties: [
          "CategoryColumn",
          "CategorySearch",
          "Path",
          "FileExtension",
          "FileName",
          "ServerRedirectedEmbedURL",
        ],
      }).then((r: SearchResults) => {
        console.log(r.PrimarySearchResults);
        let total = r.PrimarySearchResults;
        let bindSearch = "";
        total.forEach((total) => {
          total["CategorySearch"] == `${e.target["textContent"]}`
            ? (bindSearch += `<li class="doc-list"><a href="${total.Path}">${total["FileName"]}</a></li>`)
            : "";
          // bindSearch += `<li class="doc-list"><a href="${total.Path}">${total["FileName"]}</a></li>`;
        });
        console.log(bindSearch);
        let rightUL = document.querySelector(".right ul");
        rightUL.innerHTML = bindSearch;
      });
    });
  });
}
async function getSearchValue() {
  sp.search({
    Querytext:
      "Proof of Concept path:https://chandrudemo.sharepoint.com/sites/Torsion/",
    RowLimit: 100,
    // RefinementFilters: ["CategorySearch:'Proof of Concept'"],
    TrimDuplicates: true,
    SelectProperties: [
      "CategoryColumn",
      "CategorySearch",
      "Path",
      "FileExtension",
      "FileName",
      "ServerRedirectedEmbedURL",
    ],
  }).then((r: SearchResults) => {
    r.PrimarySearchResults.forEach((res) => {
      res["CategorySearch"] == "Proof of Concept" ? console.log(res) : "";
    });
  });
}
