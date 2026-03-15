/* eslint-disable */
"use strict";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import FilterAction = powerbi.FilterAction;

// Filter interfaces (from powerbi-models)
interface IFilterColumnTarget {
    table: string;
    column: string;
}

interface BasicFilter {
    $schema: string;
    target: IFilterColumnTarget;
    operator: "In" | "NotIn" | "All";
    values: (string | number | boolean)[];
    filterType: number; // 1 = BasicFilter
}

interface SlicerItem {
    value: string;
    selected: boolean;
}

interface VisualSettings {
    showOnlySelected: boolean;
    showSearch: boolean;
    multiSelect: boolean;
    fontSize: number;
    itemHeight: number;
    selectedColor: string;
    textColor: string;
    backgroundColor: string;
    borderColor: string;
}

export class SmartSlicer implements IVisual {
    private host: IVisualHost;
    private container: HTMLElement;
    private items: SlicerItem[] = [];
    private selectedValues: Set<string> = new Set();
    private searchTerm: string = "";
    private settings: VisualSettings;
    private filterTarget: IFilterColumnTarget | null = null;

    // Persistent DOM elements
    private headerEl: HTMLElement;
    private searchWrapEl: HTMLElement;
    private searchInputEl: HTMLInputElement;
    private listContainerEl: HTMLElement;
    private toggleBtnEl: HTMLButtonElement;
    private countBadgeEl: HTMLSpanElement;
    private clearBtnEl: HTMLButtonElement;

    private debounceTimer: number | null = null;
    private isInitialized: boolean = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.container = options.element as HTMLElement;
        this.settings = this.defaultSettings();
        this.buildLayout();
    }

    private defaultSettings(): VisualSettings {
        return {
            showOnlySelected: false,
            showSearch: true,
            multiSelect: true,
            fontSize: 13,
            itemHeight: 28,
            selectedColor: "#00205B",    // Alstom Blue
            textColor: "#333333",
            backgroundColor: "#FFFFFF",
            borderColor: "#CCCCCC"
        };
    }

    private buildLayout(): void {
        const s = this.settings;
        this.container.innerHTML = "";
        this.container.style.cssText = `
            background: ${s.backgroundColor};
            font-family: Segoe UI, sans-serif;
            font-size: ${s.fontSize}px;
            color: ${s.textColor};
            display: flex;
            flex-direction: column;
            height: 100%;
            box-sizing: border-box;
            overflow: hidden;
        `;

        // ── Header bar ──────────────────────────────────────
        this.headerEl = document.createElement("div");
        this.headerEl.style.cssText = `
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 4px 8px;
            flex-shrink: 0;
            gap: 6px;
        `;

        // Toggle button
        this.toggleBtnEl = document.createElement("button");
        this.toggleBtnEl.style.cssText = `
            border: none;
            background: transparent;
            color: ${s.textColor};
            border-radius: 3px;
            padding: 2px 7px;
            cursor: pointer;
            font-size: ${s.fontSize - 1}px;
            white-space: nowrap;
            flex-shrink: 0;
        `;
        this.toggleBtnEl.textContent = "☆ Selected";
        this.toggleBtnEl.addEventListener("click", () => {
            this.settings.showOnlySelected = !this.settings.showOnlySelected;
            this.persistShowOnlySelected();
            this.updateToggleButton();
            this.renderList();
        });

        // Count badge
        this.countBadgeEl = document.createElement("span");
        this.countBadgeEl.style.cssText = `
            font-size: ${s.fontSize - 2}px;
            color: #999;
            flex-shrink: 0;
        `;
        this.countBadgeEl.textContent = "None";

        // Clear button
        this.clearBtnEl = document.createElement("button");
        this.clearBtnEl.textContent = "✕ Clear";
        this.clearBtnEl.style.cssText = `
            border: none;
            background: transparent;
            color: #bbb;
            cursor: default;
            font-size: ${s.fontSize - 1}px;
            padding: 2px 4px;
            flex-shrink: 0;
        `;
        this.clearBtnEl.addEventListener("click", () => this.clearAll());

        this.headerEl.appendChild(this.toggleBtnEl);
        this.headerEl.appendChild(this.countBadgeEl);
        this.headerEl.appendChild(this.clearBtnEl);
        this.container.appendChild(this.headerEl);

        // ── Search box ───────────────────────────────────────
        this.searchWrapEl = document.createElement("div");
        this.searchWrapEl.style.cssText = `
            padding: 5px 8px;
            flex-shrink: 0;
            position: relative;
        `;

        const searchIcon = document.createElement("span");
        searchIcon.textContent = "🔍";
        searchIcon.style.cssText = `
            position: absolute;
            left: 14px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 12px;
            pointer-events: none;
        `;

        this.searchInputEl = document.createElement("input");
        this.searchInputEl.type = "text";
        this.searchInputEl.placeholder = "Search...";
        this.searchInputEl.style.cssText = `
            width: 100%;
            padding: 4px 6px 4px 24px;
            border: none;
            border-radius: 3px;
            font-size: ${s.fontSize - 1}px;
            box-sizing: border-box;
            outline: none;
            color: ${s.textColor};
            background: #f5f5f5;
        `;

        // Debounced search - only update list, not whole DOM
        this.searchInputEl.addEventListener("input", (e) => {
            this.searchTerm = (e.target as HTMLInputElement).value;
            if (this.debounceTimer) {
                clearTimeout(this.debounceTimer);
            }
            this.debounceTimer = window.setTimeout(() => {
                this.renderList();
            }, 150);
        });

        this.searchInputEl.addEventListener("focus", () => {
            this.searchInputEl.style.background = "#e8e8e8";
        });
        this.searchInputEl.addEventListener("blur", () => {
            this.searchInputEl.style.background = "#f5f5f5";
        });

        this.searchWrapEl.appendChild(searchIcon);
        this.searchWrapEl.appendChild(this.searchInputEl);
        this.container.appendChild(this.searchWrapEl);

        // ── Item list container ────────────────────────────────────
        this.listContainerEl = document.createElement("div");
        this.listContainerEl.style.cssText = `
            overflow-y: auto;
            flex: 1;
            padding: 2px 0;
        `;
        this.container.appendChild(this.listContainerEl);

        this.isInitialized = true;
    }

    private parseSettings(dataView: DataView): void {
        if (!dataView || !dataView.metadata) return;
        const objects = dataView.metadata.objects;
        if (!objects) return;

        const ss = objects["slicerSettings"] as any;
        const ap = objects["appearance"] as any;

        if (ss) {
            if (ss["showOnlySelected"] !== undefined) this.settings.showOnlySelected = ss["showOnlySelected"] as boolean;
            if (ss["showSearch"] !== undefined) this.settings.showSearch = ss["showSearch"] as boolean;
            if (ss["multiSelect"] !== undefined) this.settings.multiSelect = ss["multiSelect"] as boolean;
        }
        if (ap) {
            if (ap["fontSize"]) this.settings.fontSize = ap["fontSize"] as number;
            if (ap["itemHeight"]) this.settings.itemHeight = ap["itemHeight"] as number;
            if (ap["selectedColor"]?.solid?.color) this.settings.selectedColor = ap["selectedColor"].solid.color;
            if (ap["textColor"]?.solid?.color) this.settings.textColor = ap["textColor"].solid.color;
            if (ap["backgroundColor"]?.solid?.color) this.settings.backgroundColor = ap["backgroundColor"].solid.color;
            if (ap["borderColor"]?.solid?.color) this.settings.borderColor = ap["borderColor"].solid.color;
        }
    }

    public update(options: VisualUpdateOptions): void {
        const dataView = options.dataViews?.[0];
        if (!dataView) {
            this.items = [];
            this.filterTarget = null;
            this.renderList();
            return;
        }

        this.parseSettings(dataView);

        // Update search box visibility
        this.searchWrapEl.style.display = this.settings.showSearch ? "block" : "none";

        // Always update toggle button to reflect persisted state
        this.updateToggleButton();

        const categorical = dataView.categorical;
        if (!categorical?.categories?.length) {
            this.items = [];
            this.filterTarget = null;
            this.renderList();
            return;
        }

        const category = categorical.categories[0];
        const source = category.source;
        const values = category.values as (string | number | boolean | null)[];

        // Store the filter target (table + column)
        this.filterTarget = {
            table: source.queryName.split('.')[0],
            column: source.displayName
        };

        // Restore selections from applied filters (persists across refreshes)
        this.restoreSelectionsFromFilter(options);

        // Build items, preserving selections
        this.items = values.map((val) => {
            const strVal = val == null ? "(Blank)" : String(val);
            return {
                value: strVal,
                selected: this.selectedValues.has(strVal)
            };
        });

        // Remove duplicates
        const seen = new Set<string>();
        this.items = this.items.filter(item => {
            if (seen.has(item.value)) return false;
            seen.add(item.value);
            return true;
        });

        this.updateHeader();
        this.renderList();
    }

    private restoreSelectionsFromFilter(options: VisualUpdateOptions): void {
        // Read applied filters from Power BI to restore selection state
        const jsonFilters = options.jsonFilters;
        if (!jsonFilters || jsonFilters.length === 0) {
            // No filters applied - clear selections
            this.selectedValues.clear();
            return;
        }

        // Find our filter (BasicFilter with matching target)
        for (const filter of jsonFilters) {
            const basicFilter = filter as BasicFilter;
            if (basicFilter.filterType === 1 &&
                basicFilter.operator === "In" &&
                basicFilter.values &&
                basicFilter.target) {

                // Check if this filter targets our column
                if (this.filterTarget &&
                    basicFilter.target.table === this.filterTarget.table &&
                    basicFilter.target.column === this.filterTarget.column) {

                    // Restore selected values from filter
                    this.selectedValues.clear();
                    for (const val of basicFilter.values) {
                        this.selectedValues.add(String(val));
                    }
                    return;
                }
            }
        }
    }

    private persistShowOnlySelected(): void {
        // Persist the showOnlySelected setting to Power BI so it survives refreshes
        const instance: powerbi.VisualObjectInstance = {
            objectName: "slicerSettings",
            selector: null,
            properties: {
                showOnlySelected: this.settings.showOnlySelected
            }
        };

        this.host.persistProperties({
            merge: [instance]
        });
    }

    private updateHeader(): void {
        const s = this.settings;
        const hasSelection = this.selectedValues.size > 0;

        // Update count badge
        this.countBadgeEl.textContent = hasSelection ? `${this.selectedValues.size} selected` : "None";
        this.countBadgeEl.style.color = hasSelection ? s.selectedColor : "#999";

        // Update clear button
        this.clearBtnEl.style.color = hasSelection ? "#c00" : "#bbb";
        this.clearBtnEl.style.cursor = hasSelection ? "pointer" : "default";
        this.clearBtnEl.disabled = !hasSelection;

        this.updateToggleButton();
    }

    private updateToggleButton(): void {
        const s = this.settings;
        this.toggleBtnEl.title = s.showOnlySelected ? "Show all items" : "Show only selected";
        this.toggleBtnEl.style.background = s.showOnlySelected ? s.selectedColor : "transparent";
        this.toggleBtnEl.style.color = s.showOnlySelected ? "#fff" : s.textColor;
        this.toggleBtnEl.textContent = s.showOnlySelected ? "★ Selected" : "☆ Selected";
    }

    private toggleItem(item: SlicerItem): void {
        if (!this.settings.multiSelect) {
            // Single select
            if (this.selectedValues.has(item.value)) {
                this.selectedValues.clear();
            } else {
                this.selectedValues.clear();
                this.selectedValues.add(item.value);
            }
        } else {
            // Multi select - just toggle
            if (this.selectedValues.has(item.value)) {
                this.selectedValues.delete(item.value);
            } else {
                this.selectedValues.add(item.value);
            }
        }

        this.applyFilter();
        this.updateHeader();
        this.renderList();
    }

    private applyFilter(): void {
        if (!this.filterTarget) return;

        if (this.selectedValues.size === 0) {
            // Clear filter
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        } else {
            // Apply BasicFilter with selected values
            const filter: BasicFilter = {
                $schema: "https://powerbi.com/product/schema#basic",
                target: this.filterTarget,
                operator: "In",
                values: Array.from(this.selectedValues),
                filterType: 1 // BasicFilterType
            };

            this.host.applyJsonFilter(filter, "general", "filter", FilterAction.merge);
        }
    }

    private clearAll(): void {
        this.selectedValues.clear();
        this.applyFilter();
        this.updateHeader();
        this.renderList();
    }

    private getFilteredItems(): SlicerItem[] {
        let filtered = this.items;

        // Search filter
        if (this.searchTerm.trim()) {
            const term = this.searchTerm.toLowerCase();
            filtered = filtered.filter(i => i.value.toLowerCase().includes(term));
        }

        // Show only selected filter
        if (this.settings.showOnlySelected && this.selectedValues.size > 0) {
            filtered = filtered.filter(i => this.selectedValues.has(i.value));
        }

        return filtered;
    }

    private renderList(): void {
        const s = this.settings;
        const filtered = this.getFilteredItems();

        // Clear only the list, not the whole container
        this.listContainerEl.innerHTML = "";

        if (filtered.length === 0) {
            const empty = document.createElement("div");
            empty.style.cssText = `
                padding: 12px 8px;
                color: #999;
                font-style: italic;
                font-size: ${s.fontSize - 1}px;
                text-align: center;
            `;
            empty.textContent = this.searchTerm
                ? `No results for "${this.searchTerm}"`
                : (s.showOnlySelected ? "No items selected" : "No data");
            this.listContainerEl.appendChild(empty);
            return;
        }

        filtered.forEach(item => {
            const isSelected = this.selectedValues.has(item.value);
            const row = document.createElement("div");
            row.style.cssText = `
                display: flex;
                align-items: center;
                height: ${s.itemHeight}px;
                padding: 0 8px;
                cursor: pointer;
                border-radius: 2px;
                margin: 1px 4px;
                gap: 8px;
                transition: background 0.1s;
                background: ${isSelected ? s.selectedColor + "18" : "transparent"};
                border-left: 3px solid ${isSelected ? s.selectedColor : "transparent"};
            `;

            // Checkbox
            const checkbox = document.createElement("div");
            checkbox.style.cssText = `
                width: 14px;
                height: 14px;
                border: 1.5px solid ${isSelected ? s.selectedColor : s.borderColor};
                border-radius: 2px;
                background: ${isSelected ? s.selectedColor : "transparent"};
                display: flex;
                align-items: center;
                justify-content: center;
                flex-shrink: 0;
                transition: all 0.1s;
            `;
            if (isSelected) {
                checkbox.innerHTML = `<svg width="10" height="8" viewBox="0 0 10 8"><polyline points="1,4 4,7 9,1" fill="none" stroke="white" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
            }

            // Label
            const label = document.createElement("span");
            label.style.cssText = `
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                color: ${isSelected ? s.selectedColor : s.textColor};
                font-weight: ${isSelected ? "600" : "normal"};
                flex: 1;
                font-size: ${s.fontSize}px;
            `;

            // Highlight search term
            if (this.searchTerm.trim()) {
                const term = this.searchTerm;
                const idx = item.value.toLowerCase().indexOf(term.toLowerCase());
                if (idx >= 0) {
                    label.innerHTML =
                        this.escapeHtml(item.value.slice(0, idx)) +
                        `<mark style="background:${s.selectedColor}33;color:${s.textColor};border-radius:2px;padding:0 1px">` +
                        this.escapeHtml(item.value.slice(idx, idx + term.length)) +
                        "</mark>" +
                        this.escapeHtml(item.value.slice(idx + term.length));
                } else {
                    label.textContent = item.value;
                }
            } else {
                label.textContent = item.value;
            }

            row.appendChild(checkbox);
            row.appendChild(label);

            row.addEventListener("mouseenter", () => {
                if (!isSelected) row.style.background = s.selectedColor + "0D";
            });
            row.addEventListener("mouseleave", () => {
                row.style.background = isSelected ? s.selectedColor + "18" : "transparent";
            });
            row.addEventListener("click", () => {
                this.toggleItem(item);
            });

            this.listContainerEl.appendChild(row);
        });
    }

    private escapeHtml(str: string): string {
        return str
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    public enumerateObjectInstances(options: powerbi.EnumerateVisualObjectInstancesOptions): powerbi.VisualObjectInstanceEnumeration {
        const s = this.settings;
        switch (options.objectName) {
            case "slicerSettings":
                return [{
                    objectName: options.objectName,
                    properties: {
                        showOnlySelected: s.showOnlySelected,
                        showSearch: s.showSearch,
                        multiSelect: s.multiSelect
                    },
                    selector: null
                }];
            case "appearance":
                return [{
                    objectName: options.objectName,
                    properties: {
                        fontSize: s.fontSize,
                        itemHeight: s.itemHeight,
                        selectedColor: { solid: { color: s.selectedColor } },
                        textColor: { solid: { color: s.textColor } },
                        backgroundColor: { solid: { color: s.backgroundColor } },
                        borderColor: { solid: { color: s.borderColor } }
                    },
                    selector: null
                }];
            default:
                return [];
        }
    }
}
