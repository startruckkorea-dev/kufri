import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// GitHub Pages 커스텀 도메인(kufri.startruckkorea.com) 루트 서빙이므로 base 는 '/'
export default defineConfig({
  plugins: [react()],
  base: '/',
  build: {
    outDir: 'dist',
    sourcemap: false, // 운영 번들에 소스 노출 금지
  },
});
