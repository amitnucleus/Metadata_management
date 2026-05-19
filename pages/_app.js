import "../styles/globals.css";
import { ThemeProvider, useTheme } from "../components/ThemeContext";

function Inner({ Component, pageProps }) {
  const { theme } = useTheme();
  return (
    <div style={{
      minHeight: "100vh",
      background: theme.pageBg,
      color: theme.pageText,
      fontFamily: theme.font,
      transition: "background 0.2s, color 0.2s",
    }}>
      <Component {...pageProps} />
    </div>
  );
}

export default function App(props) {
  return (
    <ThemeProvider>
      <Inner {...props} />
    </ThemeProvider>
  );
}
