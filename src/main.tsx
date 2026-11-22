import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';
import {applyDarkTheme, getInitialDarkTheme} from './lib/theme.ts';

applyDarkTheme(getInitialDarkTheme());

createRoot(document.getElementById('root')!).render(<App />);

