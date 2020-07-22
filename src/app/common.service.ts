import { Injectable } from '@angular/core';
import axios from 'axios';

@Injectable({
  providedIn: 'root'
})


export class CommonService {
    private itemId: string;

  constructor() { }

    public getItemRestId() {
        if (Office.context.mailbox.diagnostics.hostName == 'Outlook') {
            // itemId is already REST-formatted.
            this.itemId = Office.context.mailbox.item.itemId;
            return Office.context.mailbox.item.itemId;
        } else {
            // Convert to an item ID for API v2.0.
            return Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0
            );
        }
    }

    getMessages(getMessageUrl: string) {
        return axios.get(getMessageUrl)
            .then(function (response) {
                return response.data;
            })
            .catch(function (error) {
                // handle error
                return "error he";
            })
            .finally(function () {
                // always executed
            });
    }
}
