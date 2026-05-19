import { createContext, useContext, useState, useEffect } from "react";
import { THEMES, THEME_KEYS } from "./themes";

const ThemeContext = createContext(null);

export function ThemeProvider({ children }) {
  const [themeKey, setThemeKey] = useState("dark");

  useEffect(() => {
    const saved = localStorage.getItem("ui-theme");
    if (saved && THEME_KEYS.includes(saved)) setThemeKey(saved);
  }, []);

  function setTheme(key) {
    setThemeKey(key);
    localStorage.setItem("ui-theme", key);
  }

  const theme = THEMES[themeKey];

  return (
    <ThemeContext.Provider value={{ theme, themeKey, setTheme }}>
      {children}
    </ThemeContext.Provider>
  );
}

export function useTheme() {
  const ctx = useContext(ThemeContext);
  if (!ctx) throw new Error("useTheme must be used within ThemeProvider");
  return ctx;
}
