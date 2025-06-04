import * as React from "react";
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import {
  ActionButton,
  FocusZone, // Added for keyboard navigation in filmstrip
  FocusZoneDirection, // Added for FocusZone
  getTheme,
  Icon,
  Image, // For displaying images in filmstrip items
  ImageFit, // For image fitting
  Stack,
  Text, // For displaying text
} from "office-ui-fabric-react";
import { ILink } from "../models/ILink";
import { SortableContainer, SortableElement } from "react-sortable-hoc";
import { IPnPQuickLinksProps } from "./IPnPQuickLinkProps";
import { Util } from "../util/util";

// Filmstrip might have specific item size requirements
export interface IFilmstripItemStyleProps {
  width: number;
  height: number;
}

export interface IFilmstripQuickLinksProps extends IPnPQuickLinksProps {
  itemSize: IFilmstripItemStyleProps; // e.g., { width: 150, height: 100 }
  showTitle: boolean; // Option to show title below/above image
  imageHeight?: number; // Optional: specific height for the image within the item
  borderRadius: number;
}

export const FilmstripQuickLinks: React.FunctionComponent<
  IFilmstripQuickLinksProps
> = (props: React.PropsWithChildren<IFilmstripQuickLinksProps>) => {
  const {
    displayMode,
    webPartTitle,
    setWebpartTitle,
    links,
    setLinks,
    SelectedItemId,
    setSelectedItemId,
    itemSize,
    showTitle,
    imageHeight,
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
        Title: "New Item",
        IconName: "ImagePixel", // Default for filmstrip, assuming images
        Link: "https://pnp.github.io/",
        Target: "_blank",
        // Optional: Add a placeholder image URL if your ILink supports it
        // ImageUrl: "https://via.placeholder.com/150x100"
      },
    ]);
  };

  const SortableItem = SortableElement(({ link }: { link: ILink }) => (
    <Stack
      horizontalAlign="center"
      verticalAlign="center"
      styles={{
        root: {
          width: itemSize.width,
          height: itemSize.height,
          padding: 5,
          margin: "0 5px", // Horizontal margin between items
          cursor: displayMode === DisplayMode.Edit ? "grab" : "pointer",
          border:
            SelectedItemId === link.Id
              ? `3px solid ${theme.palette.themeSecondary}`
              : `1px solid ${theme.palette.neutralLight}`,
          borderRadius: borderRadius,
          overflow: "hidden", // Ensure content fits
          boxSizing: "border-box",
        },
      }}
      onClick={
        displayMode === DisplayMode.Read
          ? () => {
              if (link.Link) window.open(link.Link, link.Target || "_self");
            }
          : () => setSelectedItemId(link.Id)
      }
      tabIndex={0} // Make it focusable
      aria-label={link.Title}
    >
      {/* Assuming IconName might store an image URL or a Fabric icon name */}
      {/* You might need a more robust way to determine if it's an image URL vs icon */}
      {link.IconName &&
      (link.IconName.startsWith("http") || link.IconName.startsWith("/")) ? (
        <Image
          src={link.IconName} // Or a dedicated ImageUrl property on ILink
          alt={link.Title}
          width={itemSize.width - 10} // Adjust for padding
          height={
            imageHeight ||
            (showTitle ? itemSize.height - 30 : itemSize.height - 10)
          } // Adjust height if title is shown
          imageFit={ImageFit.cover} // Or other fit options
          styles={{
            root: { borderRadius: borderRadius > 0 ? borderRadius - 2 : 0 },
          }}
        />
      ) : (
        <Icon
          iconName={link.IconName || "ImagePixel"}
          styles={{
            root: {
              fontSize: (imageHeight || itemSize.height) * 0.5, // Adjust icon size
              marginBottom: showTitle ? 5 : 0,
            },
          }}
        />
      )}
      {showTitle && (
        <Text
          variant="small"
          styles={{
            root: {
              textAlign: "center",
              marginTop: 5,
              whiteSpace: "nowrap",
              overflow: "hidden",
              textOverflow: "ellipsis",
              width: "100%",
            },
          }}
        >
          {link.Title}
        </Text>
      )}
    </Stack>
  ));

  const SortableList = SortableContainer(({ items }: { items: ILink[] }) => (
    <FocusZone
      direction={FocusZoneDirection.horizontal}
      style={{
        overflowX: "auto",
        overflowY: "hidden",
        whiteSpace: "nowrap",
        padding: "10px 0",
      }}
    >
      <div style={{ display: "flex", alignItems: "flex-start" }}>
        {" "}
        {/* Ensure items are aligned correctly */}
        {items.map((item, index) => (
          <SortableItem
            key={`${item.Id}`}
            index={index}
            link={item}
            disabled={displayMode === DisplayMode.Read}
          />
        ))}
      </div>
    </FocusZone>
  ));

  const UpdateSortIndex = (indexToMove: number, newIndex: number): void => {
    const arr = [...links].sort((a, b) => a.SortWeight - b.SortWeight);
    const res = Util.CalculateNewSortWeight(links, newIndex, indexToMove);
    arr[indexToMove].SortWeight = res;
    setLinks(arr);
  };

  // Basic navigation (optional, could be enhanced)
  // const scrollableContainerRef = React.useRef<HTMLDivElement>(null);
  // const scroll = (direction: "left" | "right") => {
  //   if (scrollableContainerRef.current) {
  //     const scrollAmount = itemSize.width * 2; // Scroll by two items
  //     scrollableContainerRef.current.scrollBy({ left: direction === "left" ? -scrollAmount : scrollAmount, behavior: "smooth" });
  //   }
  // };

  return (
    <div style={{ marginTop: marginTop ? `${marginTop}px` : undefined }}>
      <WebPartTitle
        displayMode={displayMode}
        title={webPartTitle}
        updateProperty={setWebpartTitle}
      />

      {displayMode === DisplayMode.Edit && (
        <ActionButton iconProps={{ iconName: "Add" }} onClick={() => AddLink()}>
          Add item
        </ActionButton>
      )}

      {/* Optional: Navigation buttons */}
      {/* <Stack horizontal horizontalAlign="space-between" styles={{ root: { marginBottom: 10 } }}>
        <IconButton iconProps={{ iconName: "ChevronLeft" }} onClick={() => scroll("left")} title="Scroll Left" />
        <IconButton iconProps={{ iconName: "ChevronRight" }} onClick={() => scroll("right")} title="Scroll Right" />
      </Stack> */}

      {/* <div ref={scrollableContainerRef} style={{ overflowX: "auto", overflowY: "hidden", whiteSpace: "nowrap", padding: "10px 0" }}> */}
      <SortableList
        items={[...links].sort((a, b) => a.SortWeight - b.SortWeight)}
        axis="x" // Filmstrip is horizontal
        distance={10}
        onSortEnd={(sort) => UpdateSortIndex(sort.oldIndex, sort.newIndex)}
        // helperClass="dragging-helper-class" // Optional: for styling the dragged item
      />
      {/* </div> */}
    </div>
  );
};
