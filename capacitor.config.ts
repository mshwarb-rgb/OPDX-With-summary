import { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'org.opd.logger',
  appName: 'OPD Logger',
  webDir: 'www',
  bundledWebRuntime: false,
  android: {
    // allow cleartext if needed for local testing; remove for production
    allowMixedContent: true
  },
  server: {
    androidScheme: 'https'
  }
};

export default config;
