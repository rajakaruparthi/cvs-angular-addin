/// <reference types='office-js' />

import {Component} from '@angular/core';
import {CommonService} from './common.service';
import axios, {AxiosResponse} from 'axios';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})

export class AppComponent {
    title = 'name';
    hostname = 'hostname';
    getMessageUrl = 'messageUrl';
    itemId = 'itemId';
    accessToken = 'accessToken';
    restId = ' restId';
    messages = 'messages';
    emails: string;
    error: string;
    status: string;
    url = '';

    constructor(private commonService: CommonService) {}

    onSub() {
        Office.initialize = res => {
            this.title = Office.context.contentLanguage;
        };

        this.hostname = Office.context.mailbox.diagnostics.hostName;
        this.restId = this.commonService.getItemRestId();
        this.getMessageUrl = Office.context.mailbox.restUrl + '/v2.0/me/messages/' + this.restId;

        Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function (asyncResult: Office.AsyncResult<string>) {
            localStorage.setItem('status', asyncResult.status.toString());
            const token = asyncResult.value;
            localStorage.setItem('apiAccessToken', token);
        });

        this.useToken();
    }

    useToken(){
        this.accessToken = localStorage.getItem('apiAccessToken');
        this.status = localStorage.getItem('status');

        const instance = axios.create({
            baseURL: this.getMessageUrl,
            timeout: 1000,
            headers: {'Authorization': 'Bearer ' + localStorage.getItem('apiAccessToken')}
        });

        instance.get(this.getMessageUrl)
            .then(response => {
                this.emails = 'came';
                return response.data;
            }).catch(error => {
                this.error = error.toString();
                this.emails =   (error.response) + '-' + (error.response.status) + '-' +  (error.response.headers);
                return 'error';
        });
    }
}

