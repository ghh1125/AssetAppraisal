import { createApp } from 'vue'
import {
  Alert,
  Button,
  Card,
  Checkbox,
  DatePicker,
  Divider,
  Form,
  Input,
  InputNumber,
  Progress,
  Radio,
  Select,
  Tag,
  Upload,
} from 'ant-design-vue'
import 'ant-design-vue/dist/reset.css'
import App from './App.vue'
import './style.css'
import { i18n } from './i18n'

const app = createApp(App)

// Register only the components used by this single-page application. Registering
// the complete Ant Design Vue plugin pulls every component into the initial bundle.
;[
  Alert,
  Button,
  Card,
  Checkbox,
  DatePicker,
  Divider,
  Form,
  Input,
  InputNumber,
  Progress,
  Radio,
  Select,
  Tag,
  Upload,
].forEach(component => app.use(component))

app.use(i18n).mount('#app')
