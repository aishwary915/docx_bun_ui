/**
 * Custom Quick Insert Menu Component
 *
 * React component that renders the slash command menu with AI button
 */

import React, { useCallback, useEffect, useRef, useState } from "react";
import {
  CustomQuickInsertService,
  type IQuickInsertMenuItem,
  type QuickInsertMenu,
} from "../plugins/CustomQuickInsertPlugin";

interface CustomQuickInsertMenuProps {
  service: CustomQuickInsertService;
  visible: boolean;
  position?: { x: number; y: number };
  onSelect: (menu: IQuickInsertMenuItem) => void;
  onClose: () => void;
}

function filterMenusByKeyword(
  menus: QuickInsertMenu[],
  keyword: string
): QuickInsertMenu[] {
  return menus
    .map((menu) => ({ ...menu }))
    .filter((menu) => {
      if ("children" in menu) {
        menu.children = filterMenusByKeyword(
          menu.children!,
          keyword
        ) as IQuickInsertMenuItem[];
        return menu.children.length > 0;
      }

      const keywords = (menu as IQuickInsertMenuItem).keywords;
      if (keywords) {
        return keywords.some((word) =>
          word.toLowerCase().includes(keyword.toLowerCase())
        );
      }

      return menu.title.toLowerCase().includes(keyword.toLowerCase());
    });
}

export function CustomQuickInsertMenu({
  service,
  visible,
  position,
  onSelect,
  onClose,
}: CustomQuickInsertMenuProps) {
  const [filterKeyword, setFilterKeyword] = useState("");
  const [focusedIndex, setFocusedIndex] = useState(0);
  const [menus, setMenus] = useState<QuickInsertMenu[]>([]);
  const [filteredMenus, setFilteredMenus] = useState<QuickInsertMenu[]>([]);
  const menuRef = useRef<HTMLDivElement>(null);
  const focusedMenuRef = useRef<IQuickInsertMenuItem | null>(null);
  const menuIndexAccumulator = useRef(0);

  // Subscribe to filter keyword changes
  useEffect(() => {
    if (!service) return;

    const subscription = service.filterKeyword$.subscribe((keyword) => {
      setFilterKeyword(keyword);
    });

    return () => {
      subscription.unsubscribe();
    };
  }, [service]);

  // Load menus
  useEffect(() => {
    if (!service) return;
    const loadedMenus = service.getMenus();
    setMenus(loadedMenus);
    setFilteredMenus(loadedMenus);
  }, [service, visible]);

  // Filter menus when keyword changes
  useEffect(() => {
    const filtered = filterMenusByKeyword(menus, filterKeyword);
    setFilteredMenus(filtered);
    setFocusedIndex(0);
  }, [menus, filterKeyword]);

  // Handle keyboard navigation
  useEffect(() => {
    if (!visible) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        onClose();
        return;
      }

      if (e.key === "ArrowDown") {
        e.preventDefault();
        e.stopPropagation();
        setFocusedIndex((prev) => {
          const max = menuIndexAccumulator.current - 1;
          return prev < max ? prev + 1 : 0;
        });
        return;
      }

      if (e.key === "ArrowUp") {
        e.preventDefault();
        e.stopPropagation();
        setFocusedIndex((prev) => {
          const max = menuIndexAccumulator.current - 1;
          return prev > 0 ? prev - 1 : max;
        });
        return;
      }

      if (e.key === "Enter") {
        e.preventDefault();
        e.stopPropagation();
        if (focusedMenuRef.current) {
          onSelect(focusedMenuRef.current);
        }
        return;
      }
    };

    window.addEventListener("keydown", handleKeyDown, true);
    return () => {
      window.removeEventListener("keydown", handleKeyDown, true);
    };
  }, [visible, onClose, onSelect]);

  // Click outside to close
  useEffect(() => {
    if (!visible) return;

    const handleClickOutside = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        console.log("[CustomQuickInsertMenu] Closing due to click outside");
        onClose();
      }
    };

    // Increased delay to prevent immediate closing - longer delay for end-of-file cases
    const timeoutId = setTimeout(() => {
      document.addEventListener("mousedown", handleClickOutside);
      console.log("[CustomQuickInsertMenu] Click-outside handler attached");
    }, 300);

    return () => {
      clearTimeout(timeoutId);
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [visible, onClose]);

  const handleMenuClick = useCallback(
    (menu: IQuickInsertMenuItem) => {
      onSelect(menu);
    },
    [onSelect]
  );

  const renderMenus = (menuList: QuickInsertMenu[]): React.ReactNode[] => {
    return menuList.map((menu) => {
      if ("children" in menu) {
        return (
          <div key={menu.id} className="mb-3 first:mt-0">
            <div className="px-4 py-1.5 flex items-center gap-2">
              {menu.icon && (
                <span className="text-base opacity-60">{menu.icon}</span>
              )}
              <span className="text-xs font-semibold uppercase tracking-wider text-gray-400 dark:text-gray-500">
                {menu.title}
              </span>
            </div>
            <div className="mt-1">{renderMenus(menu.children!)}</div>
          </div>
        );
      }

      const currentMenuIndex = menuIndexAccumulator.current;
      const isFocused = focusedIndex === currentMenuIndex;

      if (isFocused) {
        focusedMenuRef.current = menu as IQuickInsertMenuItem;
      }

      menuIndexAccumulator.current++;

      const isAIItem = (menu as IQuickInsertMenuItem).isAI;

      return (
        <div
          key={menu.id}
          className={`
            group relative px-4 py-2.5 mx-2 rounded-xl cursor-pointer flex items-center gap-3
            transition-all duration-200 ease-out
            ${
              isFocused
                ? "bg-gradient-to-r from-blue-50 to-indigo-50 dark:from-blue-900/30 dark:to-indigo-900/30 shadow-sm scale-[1.02]"
                : "hover:bg-gray-50 dark:hover:bg-gray-800/50"
            }
            ${isAIItem ? "border-l-3 border-blue-500" : ""}
          `}
          onMouseEnter={() => setFocusedIndex(currentMenuIndex)}
          onClick={() => handleMenuClick(menu as IQuickInsertMenuItem)}
        >
          {/* Icon with background */}
          {menu.icon && (
            <div className={`
              w-8 h-8 rounded-lg flex items-center justify-center text-lg flex-shrink-0
              transition-all duration-200
              ${
                isFocused
                  ? "bg-white dark:bg-gray-700 shadow-sm"
                  : "bg-gray-100 dark:bg-gray-800 group-hover:bg-white dark:group-hover:bg-gray-700"
              }
              ${isAIItem ? "bg-gradient-to-br from-blue-100 to-indigo-100 dark:from-blue-900/50 dark:to-indigo-900/50" : ""}
            `}>
              {menu.icon}
            </div>
          )}
          
          {/* Title */}
          <span className={`
            text-sm font-medium flex-1 truncate
            transition-colors duration-200
            ${
              isFocused
                ? "text-gray-900 dark:text-white"
                : "text-gray-700 dark:text-gray-300"
            }
          `}>
            {menu.title}
          </span>
          
          {/* AI Badge */}
          {isAIItem && (
            <div className="flex items-center gap-1.5 px-2 py-1 bg-gradient-to-r from-blue-500 to-indigo-500 rounded-full shadow-sm">
              <span className="text-[10px] font-bold text-white tracking-wide">AI</span>
              <span className="text-white text-xs">✨</span>
            </div>
          )}
          
          {/* Hover indicator */}
          {isFocused && (
            <div className="absolute right-3 top-1/2 -translate-y-1/2">
              <div className="w-1.5 h-1.5 rounded-full bg-blue-500 animate-pulse"></div>
            </div>
          )}
        </div>
      );
    });
  };

  if (!visible) return null;

  // Reset menu index accumulator before rendering
  menuIndexAccumulator.current = 0;

  const hasMenus = filteredMenus.length > 0;

  console.log(
    "[CustomQuickInsertMenu] Rendering menu at position:",
    position,
    "visible:",
    visible
  );

  return (
    <div
      ref={menuRef}
      className="fixed z-[9999] bg-white/95 dark:bg-gray-900/95 backdrop-blur-xl rounded-2xl shadow-2xl border border-gray-200/50 dark:border-gray-700/50 overflow-hidden animate-in fade-in slide-in-from-bottom-2 duration-200"
      style={{
        top: position?.y ? `${position.y}px` : "50%",
        left: position?.x ? `${position.x}px` : "50%",
        transform: !position ? "translate(-50%, -50%)" : undefined,
        minWidth: "280px",
        maxWidth: "360px",
        maxHeight: "480px",
        boxShadow: "0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04), 0 0 0 1px rgba(0, 0, 0, 0.05)",
      }}
    >
      {/* Header with gradient */}
      <div className="px-4 py-3 bg-gradient-to-r from-blue-50 to-indigo-50 dark:from-gray-800 dark:to-gray-800 border-b border-gray-200/50 dark:border-gray-700/50">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-2 h-2 rounded-full bg-blue-500 animate-pulse"></div>
            <span className="text-sm font-semibold text-gray-700 dark:text-gray-300">
              Quick Insert
            </span>
          </div>
          {filterKeyword && (
            <div className="flex items-center gap-1.5 px-2.5 py-1 bg-blue-100 dark:bg-blue-900/30 rounded-full">
              <span className="text-xs font-medium text-blue-600 dark:text-blue-400">
                {filterKeyword}
              </span>
            </div>
          )}
        </div>
      </div>

      {/* Menu items with custom scrollbar */}
      <div className="py-2 max-h-[360px] overflow-y-auto scrollbar-thin scrollbar-thumb-gray-300 dark:scrollbar-thumb-gray-700 scrollbar-track-transparent">
        {hasMenus ? (
          renderMenus(filteredMenus)
        ) : (
          <div className="px-4 py-8 text-center">
            <div className="text-4xl mb-2">🔍</div>
            <p className="text-sm text-gray-500 dark:text-gray-400">
              No matching commands
            </p>
            <p className="text-xs text-gray-400 dark:text-gray-500 mt-1">
              Try a different search term
            </p>
          </div>
        )}
      </div>

      {/* Footer with better styling */}
      <div className="px-4 py-2.5 bg-gray-50 dark:bg-gray-800/50 border-t border-gray-200/50 dark:border-gray-700/50">
        <div className="flex items-center gap-4 text-xs text-gray-500 dark:text-gray-400">
          <div className="flex items-center gap-1.5">
            <kbd className="px-1.5 py-0.5 bg-white dark:bg-gray-700 border border-gray-300 dark:border-gray-600 rounded text-[10px] font-mono shadow-sm">↑↓</kbd>
            <span>Navigate</span>
          </div>
          <div className="flex items-center gap-1.5">
            <kbd className="px-1.5 py-0.5 bg-white dark:bg-gray-700 border border-gray-300 dark:border-gray-600 rounded text-[10px] font-mono shadow-sm">↵</kbd>
            <span>Select</span>
          </div>
          <div className="flex items-center gap-1.5">
            <kbd className="px-1.5 py-0.5 bg-white dark:bg-gray-700 border border-gray-300 dark:border-gray-600 rounded text-[10px] font-mono shadow-sm">Esc</kbd>
            <span>Close</span>
          </div>
        </div>
      </div>
    </div>
  );
}

export default CustomQuickInsertMenu;
