// @ts-check
import { defineConfig } from 'astro/config';

import tailwind from '@astrojs/tailwind';

// https://astro.build/config
export default defineConfig({
  site: 'https://TU-USUARIO-DE-GITHUB.github.io',
  base: '/NOMBRE-DE-TU-REPOSITORIO',
  integrations: [tailwind()]
});