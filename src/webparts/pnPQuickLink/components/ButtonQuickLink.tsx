import * as React from "react";
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import {
  ActionButton,
  DefaultButton, // Changed from CompoundButton for a simpler button style
  getTheme,
  Icon,
} from "office-ui-fabric-react";
import { ILink } from "../models/ILink";
import { SortableContainer, SortableElement } from "react-sortable-hoc";
// import { TileSize } from "../models/enums"; // Not needed for Button layout
import { IPnPQuickLinksProps } from "./IPnPQuickLinkProps";
import { Util } from "../util/util";

export interface IButtonQuickLinksProps extends IPnPQuickLinksProps {
  // size: TileSize; // Removed, button size might be handled differently or fixed
  hideText: boolean;
  borderRadius: number; // May or may not be applicable to buttons, keeping for now
}

export const ButtonQuickLinks: React.FunctionComponent<
  IButtonQuickLinksProps
> = (props: React.PropsWithChildren<IButtonQuickLinksProps>) => {
  const {
    displayMode,
    webPartTitle,
    setWebpartTitle,
    links,
    setLinks,
    SelectedItemId,
    setSelectedItemId,
    hideText,
    // size, // Removed
    marginTop,
    borderRadius,
  } = props;
  const theme = getTheme();

  const AddLink = (): void => {
    setLinks([
      ...links,
      {
        Id: Util.GenerateId(),
        SortWeight: Util.CalculateNewSortWeight(links, links.length),
        Title: "New link",
        IconName: "Link", // Changed default icon
        Link: "https://pnp.github.io/",
        Target: "_blank",
      },
    ]);
  };

  const SortableItem = SortableElement(({ link }: { link: ILink }) => (
    <DefaultButton // Changed from CompoundButton
      text={!hideText ? link.Title : undefined} // Hide text if hideText is true
      iconProps={!hideText ? { iconName: link.IconName } : undefined} // Show icon only if text is hidden or always show?
      styles={{
        root: {
          border:
            SelectedItemId === link.Id
              ? `3px solid ${theme.palette.themeSecondary}`
              : `1px solid ${theme.palette.neutralLight}`, // Standard border for buttons
          // width: "auto", // Buttons typically size to content
          minHeight: 40, // Adjust as needed
          borderRadius: borderRadius,
          margin: "5px", // Add some margin between buttons
          padding: "10px 15px",
        },
        textContainer: {
          // May not be needed or needs adjustment for DefaultButton
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
        },
        icon: {
          // Style for icon if shown without text
          fontSize: 20, // Example size
          marginRight: hideText ? 0 : 5, // Adjust margin if text is present
        },
      }}
      href={displayMode === DisplayMode.Read ? link.Link : undefined}
      target={displayMode === DisplayMode.Read ? link.Target : undefined}
      onClick={
        displayMode === DisplayMode.Edit
          ? () => setSelectedItemId(link.Id)
          : undefined
      }
      onRenderIcon={
        hideText
          ? () => (
              <Icon
                iconName={link.IconName}
                styles={{ root: { fontSize: 20 } }}
              />
            )
          : undefined
      }
      onRenderText={hideText ? () => null : undefined}
    />
  ));

  const SortableList = SortableContainer(({ items }: { items: ILink[] }) => (
    <div
      style={{
        gap: 10,
        display: "flex",
        flexWrap: "wrap",
        alignItems: "flex-start",
      }}
    >
      {items.map((item, index) => (
        <SortableItem
          key={`${item.Id}`}
          index={index}
          link={item}
          disabled={displayMode === DisplayMode.Read}
        />
      ))}
    </div>
  ));

  const UpdateSortIndex = (indexToMove: number, newIndex: number): void => {
    const arr = [...links].sort((a, b) => a.SortWeight - b.SortWeight);
    const res = Util.CalculateNewSortWeight(links, newIndex, indexToMove);
    arr[indexToMove].SortWeight = res;
    setLinks(arr);
  };

  return (
    <div style={{ marginTop: marginTop ? `${marginTop}px` : undefined }}>
      <WebPartTitle
        displayMode={displayMode}
        title={webPartTitle}
        updateProperty={setWebpartTitle}
      />

      {displayMode === DisplayMode.Edit && (
        <ActionButton iconProps={{ iconName: "Add" }} onClick={() => AddLink()}>
          Add link
        </ActionButton>
      )}

      <SortableList
        items={[...links].sort((a, b) => a.SortWeight - b.SortWeight)}
        axis="xy" // For buttons, 'xy' or 'x' might be appropriate
        distance={10} // Adjust distance for drag sensitivity
        onSortEnd={(sort) => UpdateSortIndex(sort.oldIndex, sort.newIndex)}
      />
    </div>
  );
};
