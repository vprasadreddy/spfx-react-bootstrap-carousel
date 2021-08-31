import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext
} from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp/presets/all";
import * as strings from "SpfxReactBootstrapCarouselWebPartStrings";
import SpfxReactBootstrapCarousel from "./components/SpfxReactBootstrapCarousel";
import { ISpfxReactBootstrapCarouselProps } from "./components/ISpfxReactBootstrapCarouselProps";

export interface ISpfxReactBootstrapCarouselWebPartProps {
  description: string;
  context: WebPartContext;
  pictureLibraryDropDown: string;
}

export default class SpfxReactBootstrapCarouselWebPart extends BaseClientSideWebPart<ISpfxReactBootstrapCarouselWebPartProps> {
  private libraries: IPropertyPaneDropdownOption[] = [];

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  /* to add apply button in property pane */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public getAllLibraries(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve, reject) => {
      sp.web.lists
        .select("Title")
        .filter(`(BaseType eq 1) and (Hidden eq false)`)
        .get()
        .then((result) => {
          let libraries = [];
          result.map((result) => {
            libraries.push({
              key: result.Title,
              text: result.Title
            });
          });
          resolve(libraries);
        })
        .catch((error) => {
          console.log(error);
          reject("error while fetching lists");
        });
    });
  }
  public render(): void {
    const element: React.ReactElement<ISpfxReactBootstrapCarouselProps> =
      React.createElement(SpfxReactBootstrapCarousel, {
        description: this.properties.description,
        context: this.context,
        pictureLibraryDropDown: this.properties.pictureLibraryDropDown
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown("pictureLibraryDropDown", {
                  label: strings.PictureLibraryDropDownLabel,
                  options: this.libraries,
                  selectedKey: ""
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.getAllLibraries().then((result: IPropertyPaneDropdownOption[]) => {
      let defaultValue: IPropertyPaneDropdownOption[] = [
        {
          key: "",
          text: ""
        }
      ];
      this.libraries = [...defaultValue, ...result];
      console.log(this.libraries);
      //to refresh property pane after getting dropdown options
      this.context.propertyPane.refresh();
    });
  }
}
