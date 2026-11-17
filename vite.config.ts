import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const oauthHeaders = {
    'Cross-Origin-Opener-Policy': 'same-origin-allow-popups',
};

export default defineConfig({
    base: '/Mock_Test_Practice/',
    plugins: [react()],
    server: { headers: oauthHeaders },
    preview: { headers: oauthHeaders },
});