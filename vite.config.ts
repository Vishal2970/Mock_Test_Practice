import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
export default defineConfig(
    {
        base: '/Mock_Test_Practice/',
        plugins: [react()],
    }
);