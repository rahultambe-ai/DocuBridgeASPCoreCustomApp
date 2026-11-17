import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import App from './App';

// Icons used by Fluent UI controls. In SharePoint the page did this for the web part.
initializeIcons();

ReactDOM.render(<App />, document.getElementById('root'));
