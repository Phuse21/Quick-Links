import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConditionalGroup,
  IPropertyPaneConfiguration,
  IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneSlider, // Added PropertyPaneSlider
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { PropertyFieldIconPicker } from "@pnp/spfx-property-controls/lib/PropertyFieldIconPicker";
import { ILink } from "./models/ILink";
import { LayoutType, TileSize, LinkAlignment } from "./models/enums"; // Added LinkAlignment
import { IPnPQuickLinksProps } from "./components/IPnPQuickLinkProps";
import {
  ITilesQuickLinksProps,
  TilesQuickLinks,
} from "./components/TilesQuickLink";
// Removed Button and Filmstrip imports

export interface IPnPQuickLinksWebPartProps {
  webpartTitle: string;
  links: ILink[];
  type: LayoutType; // Will always be Tiles
  tile: {
    size: TileSize;
    hideText: boolean;
    borderRadius: number;
    linkAlignment?: LinkAlignment; // Added linkAlignment property
  };
  // Removed button and filmstrip properties
  marginTop?: number;
}

export default class PnPQuickLinksWebPart extends BaseClientSideWebPart<IPnPQuickLinksWebPartProps> {
  private SelectedItemId: string = "";

  public render(): void {
    const SharedProps: IPnPQuickLinksProps = {
      webPartTitle: this.properties.webpartTitle,
      setWebpartTitle: (val: string) => {
        this.properties.webpartTitle = val;
      },
      links: this.properties.links,
      setLinks: (val: ILink[]) => {
        this.properties.links = val;
        this.onPropertyPaneFieldChanged("links", null, val);
        this.render();
      },

      SelectedItemId: this.SelectedItemId,
      setSelectedItemId: (id: string) => {
        this.SelectedItemId = id;
        if (this.context.propertyPane.isPropertyPaneOpen()) {
          this.context.propertyPane.refresh();
        } else {
          this.context.propertyPane.open();
        }
        this.render();
      },

      displayMode: this.displayMode,
      marginTop: this.properties.marginTop,
    } as IPnPQuickLinksProps;

    // Ensure the type is always Tiles
    if (this.properties.type !== LayoutType.Tiles) {
      this.properties.type = LayoutType.Tiles;
    }

    const element: React.ReactElement = React.createElement(TilesQuickLinks, {
      ...SharedProps,
      hideText: this.properties.tile.hideText,
      size: this.properties.tile.size,
      borderRadius: this.properties.tile.borderRadius,
      linkAlignment: this.properties.tile.linkAlignment || LinkAlignment.Left, // Pass linkAlignment, default to Left
    } as ITilesQuickLinksProps);

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (this.SelectedItemId) {
      const index = this.properties.links
        .map((x) => x.Id)
        .indexOf(this.SelectedItemId);
      return {
        pages: [
          {
            header: {
              description: "Edit link",
            },
            groups: [
              {
                groupFields: [
                  PropertyPaneTextField(`links[${index}].Title`, {
                    label: "Title",
                  }),
                  PropertyPaneTextField(`links[${index}].Link`, {
                    label: "Url",
                  }),
                  PropertyFieldIconPicker(`links[${index}].IconName`, {
                    key: `links[${index}].IconName`,
                    // eslint-disable-next-line no-return-assign
                    onSave: (name) =>
                      (this.properties.links[index].IconName = name),
                    onPropertyChange: this.onPropertyPaneFieldChanged,
                    properties: this.properties,
                    label: "Icon",
                    buttonLabel: "Choose",
                    currentIcon: this.properties.links[index].IconName,
                  }),
                  PropertyPaneChoiceGroup(`links[${index}].Target`, {
                    label: "Open",
                    options: [
                      { key: "_self", text: "In this tab" },
                      { key: "_blank", text: "In new tab" },
                    ],
                  }),
                  PropertyPaneButton("", {
                    text: "Delete",
                    icon: "Delete",
                    onClick: () => {
                      this.properties.links = this.properties.links.filter(
                        (x) => x.Id !== this.SelectedItemId
                      );
                      this.onPropertyPaneFieldChanged(
                        "links",
                        null,
                        this.properties.links
                      );
                      this.SelectedItemId = "";
                      this.context.propertyPane.refresh();
                      this.render();
                    },
                  }),
                  PropertyPaneButton("", {
                    text: "Close",
                    onClick: () => {
                      this.SelectedItemId = "";
                      this.context.propertyPane.refresh();
                      this.render();
                    },
                  }),
                ],
              },
            ],
          },
        ],
      };
    }

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "General Settings", // Renamed
              groupFields: [
                PropertyPaneSlider("marginTop", {
                  label: "Top Margin (px)",
                  min: -100,
                  max: 30,
                  step: 1,
                  showValue: true,
                  value: this.properties.marginTop,
                }),
                // PropertyPaneChoiceGroup for "type" is removed.
              ],
            },
            // TileLayoutFields are now directly part of the groups
            ...this.TileLayoutFields(),
            {
              groupName: "Filter", // This group remains
              groupFields: [
                PropertyPaneToggle("_", {
                  label: "Enable audience targeting",
                  disabled: true,
                }),
                PropertyPaneLabel("", {
                  text: "Audience targeting is not yet implemented, feel free to do so",
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private TileLayoutFields(): (
    | IPropertyPaneGroup
    | IPropertyPaneConditionalGroup
  )[] {
    return [
      {
        groupName: "Tile Layout Options", // Added group name for clarity
        groupFields: [
          PropertyPaneChoiceGroup("tile.size", {
            label: "Icon size",
            options: [
              { key: TileSize.Small, text: "Small" },
              { key: TileSize.Medium, text: "Medium" },
              { key: TileSize.Large, text: "Large" },
              { key: TileSize.XLarge, text: "Extra large" },
              { key: TileSize.FillSpace, text: "Fill space" },
            ],
          }),
          PropertyPaneToggle("tile.hideText", {
            label: "Show only icon",
            disabled: this.properties.tile.size > TileSize.Large,
          }),
          PropertyPaneSlider("tile.borderRadius", {
            label: "Border Radius (px)",
            min: 0,
            max: 100,
            step: 1,
            showValue: true,
            value: this.properties.tile.borderRadius,
          }),
          PropertyPaneChoiceGroup("tile.linkAlignment", {
            label: "Link Alignment",
            options: [
              {
                key: LinkAlignment.Left,
                text: "Left",
                iconProps: { officeFabricIconFontName: "AlignLeft" },
              },
              {
                key: LinkAlignment.Center,
                text: "Center",
                iconProps: { officeFabricIconFontName: "AlignCenter" },
              },
              {
                key: LinkAlignment.Right,
                text: "Right",
                iconProps: { officeFabricIconFontName: "AlignRight" },
              },
            ],
            // selectedKey is removed as it's managed by the property value itself
          }),
        ],
      },
    ];
  }
  // Unused ButtonLayoutFields and FilmstripLayoutFields are removed.
}
