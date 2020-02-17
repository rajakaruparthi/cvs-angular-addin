import { Component } from '@angular/core';
import {CommonService} from './common.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'cvs-angular-addin';
  constructor(private commonService: CommonService){}

  onSub() {
    this.title = this.commonService.getMessage();

    if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
      // Perform actions.
    }
    else {
      // Provide alternate flow/logic.
    }
  }
}
