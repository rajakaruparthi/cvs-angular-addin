import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}
// Office.initialize = temp => {
// };
Office.initialize = reason =>{
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.error(err));
};
/*(async () => {
    await Office.onReady();
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.error(err));
})();
Office.initialize = reason =>{
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.error(err));
};

Office.onReady().then(function () {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.error(err));
}); */