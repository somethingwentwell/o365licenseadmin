import { defineStore } from 'pinia';
import { Loading } from 'quasar'

type user = {
  id: string,
  displayName: string,
  givenName: string,
  jobTitle: string,
  mail: string,
  officeLocation: string,
  preferredLanguage: string,
  surname: string,
  userPrincipalName: string
}

type customUser = {
  id: string,
  displayName: string,
  jobTitle: string,
  mail: string,
  officeLocation: string,
  assignedLicenses: [
    {
        disabledPlans: Array<string>,
        skuId: string
    }
  ]
}

type currentUser = {
  id: string,
  displayName: string
}

type sku = {
  id: string,
  capabilityStatus: Array<string>,
  consumedUnits: string,
  skuId: string,
  skuPartNumber: string,
  appliesTo: string,
  prepaidUnits: {
    enabled: number,
    suspended: number,
    warning: number
  },
  servicePlans: [
    {
      servicePlanId: string,
      servicePlanName: string,
      provisioningStatus: string,
      appliesTo: string
    }
  ]
}

type log = {
  PartitionKey: string,
  RowKey: string,
  Timestamp: string,
  success: boolean,
  RequestBody: string,
  Result: string
}

export const useO365Store = defineStore('o365', {
  state: () => ({
    storage: {
      accountName: 'mvpkso3658b7d',
      sasToken: 'sv=2021-06-08&ss=bfqt&srt=sco&sp=rwdlacupiyx&se=2022-11-29T19:40:14Z&st=2022-10-18T11:40:14Z&spr=https&sig=We%2FifJ8ZV%2BDKGbnzHdF53pML5RFQVhbY6MaqpPdJvf0%3D',
      authUrl: 'https://kaishingo365.azurewebsites.net/.auth/login/aad?post_login_redirect_url=/api/GetAccessTokenAuth?url='
    },
    users: [] as user[],
    customUsers: [] as customUser[],
    skus: [] as sku[],
    currentUser: {} as currentUser,
    assignedLicense: [
      {
      'id': '',
      'skuId': '',
      'skuPartNumber': '',
      'sercicePlans': [
        {
          'servicePlanId': '',
          'servicePlanName': '',
          'provisioningStatus': '',
          'appliesto': ''
        }
      ]
    }],
    logs: [] as log[],
    config: {
      PartitionKey: '',
      RowKey: '',
      Timestamp: '',
      email: '',
      deactivate_period: 0,
      delete_period: 0
    }
  }),

  getters: {
    getUsers: (state) => {
      return state.users;
    },
    getSkus: (state) => {
      return state.skus;
    }
  },

  actions: {
    // async getGraphToken() {
    //   Loading.show();
    //   const response = await fetch('https://kaishingo365.azurewebsites.net/api/GetAccessToken');
    //   const result = await response.json();
    //   // console.log(result.access_token);
    //   Loading.hide()
    //   return result.access_token;
    // },
    async getGraphUsers() {
      Loading.show();
      const myHeaders = new Headers();
      // const token = await this.getGraphToken();
      const token = localStorage.getItem('token');
      myHeaders.append('Authorization', `Bearer ${token}`);
      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
      };
      const response = await fetch('https://graph.microsoft.com/v1.0/users', requestOptions);
      const result = await response.json();

      this.users = result.value;
      Loading.hide();
      return;
    },
    async getGraphCustomUsers() {
      Loading.show();
      const myHeaders = new Headers();
      // const token = await this.getGraphToken();
      const token = localStorage.getItem('token');
      myHeaders.append('Authorization', `Bearer ${token}`);
      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
      };
      const response = await fetch('https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,jobTitle,officeLocation,assignedLicenses', requestOptions);
      const result = await response.json();

      this.customUsers = result.value;
      Loading.hide();
      return;
    },
    async getGraphSkus() {
      Loading.show();
      const myHeaders = new Headers();
      // const token = await this.getGraphToken();
      const token = localStorage.getItem('token');
      myHeaders.append('Authorization', `Bearer ${token}`);
      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
      };
      const response = await fetch('https://graph.microsoft.com/v1.0/subscribedSkus', requestOptions);
      const result = await response.json();
      this.skus = result.value;
      Loading.hide();
      return;
    },
    async getLicenseDetails(userId: string) {
      Loading.show();
      const myHeaders = new Headers();
      // const token = await this.getGraphToken();
      const token = localStorage.getItem('token');
      myHeaders.append('Authorization', `Bearer ${token}`);
      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
      };
      const response = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/licenseDetails`, requestOptions);
      const result = await response.json();
      this.assignedLicense = result.value;
      Loading.hide();
      return;
    },
    async removeUserLicense(userId: string, skuId: string) {
      Loading.show();
      const myHeaders = new Headers();
      // const token = await this.getGraphToken();
      const token = localStorage.getItem('token');
      myHeaders.append('Authorization', `Bearer ${token}`);
      myHeaders.append('Content-Type', 'application/json');
      const raw = JSON.stringify({
        'addLicenses': [],
        'removeLicenses': [
          skuId
        ]
      });
      const requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: raw
      };
      const response = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/assignLicense`, requestOptions);
      const message = await response.text();
      if (!response.ok) {
        this.logApiResult(raw, message, false, 'removeUserLicense');
        alert(message);
      }
      else {
        this.logApiResult(raw, message, true, 'removeUserLicense');
      }
      Loading.hide();
      return;
    },
    async addUserLicense(userId: string, skuIds: string[]) {
      Loading.show();
      const myHeaders = new Headers();
      // const token = await this.getGraphToken();
      const token = localStorage.getItem('token');
      myHeaders.append('Authorization', `Bearer ${token}`);
      myHeaders.append('Content-Type', 'application/json');
      const addedSkus = [];
      skuIds.forEach(element => {
        addedSkus.push({
          'skuId': element
        })
      });
      const raw = JSON.stringify({
        'addLicenses': addedSkus,
        'removeLicenses': []
      });
      const requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: raw
      };
      const response = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/assignLicense`, requestOptions);
      const message = await response.text();
      if (!response.ok) {
        this.logApiResult(raw, message, false, 'addUserLicense');
        alert(message);
      }
      else {
        this.logApiResult(raw, message, true, 'addUserLicense');
      }
      Loading.hide();
      return;
    },
    async logApiResult(requestBody: string, result: string, success: boolean, partition: string) {
      Loading.show();
      const myHeaders = new Headers();
      myHeaders.append('Content-Type', 'application/json');

      const raw = JSON.stringify({
        'RowKey': partition + Date.now(),
        'PartitionKey': partition,
        'success': success,
        'RequestBody': requestBody,
        'Result': result
      });
      const requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: raw
      };
      await fetch(`https://${this.storage.accountName}.table.core.windows.net/logs?${this.storage.sasToken}`, requestOptions);
      Loading.hide();
      return;
    },
    async getLogs() {
      Loading.show();
      const myHeaders = new Headers();
      myHeaders.append('Accept', 'application/json');

      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
      };
      const response = await fetch(`https://${this.storage.accountName}.table.core.windows.net/logs?$top=50&${this.storage.sasToken}`, requestOptions);
      const message = await response.json();
      this.logs = message.value;
      Loading.hide();
      return;
    },
    async getConfig() {
      Loading.show();
      const myHeaders = new Headers();
      myHeaders.append('Accept', 'application/json');

      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
      };
      const response = await fetch(`https://${this.storage.accountName}.table.core.windows.net/config?${this.storage.sasToken}`, requestOptions);
      const message = await response.json();
      this.config = message.value[0];
      Loading.hide();
      return;
    },
    async updateConfig() {
      Loading.show();
      const myHeaders = new Headers();
      myHeaders.append('Accept', 'application/json');
      myHeaders.append('Content-Type', 'application/json');

      const raw = JSON.stringify({
        'email': this.config.email,
        'deactivate_period': this.config.deactivate_period,
        'delete_period': this.config.delete_period
      });

      const requestOptions = {
        method: 'PUT',
        headers: myHeaders,
        body: raw
      };

      const key = 'automation';
      await fetch(`https://${this.storage.accountName}.table.core.windows.net/config(PartitionKey='${key}', RowKey='${key}')?${this.storage.sasToken}`, requestOptions);
      Loading.hide();
      return;
    },
    auth() {
      const urlParams = new URLSearchParams(window.location.search);

      const redirectAuth = () => {
        console.log(this.storage.authUrl + window.location.href)
        localStorage.setItem('href', window.location.href);
        window.location.replace(this.storage.authUrl + window.location.href);
      }

      const setToken = () => {
        const data = JSON.parse(atob(urlParams.get('data')));
        localStorage.setItem('token', data.access_token);
        localStorage.setItem('date', Date.now().toString());
        localStorage.setItem('expiry', data.expires_in);
        window.location.replace(localStorage.getItem('href'));
        localStorage.removeItem('href');
      }


      if (localStorage.getItem('token') == null) {
        if (urlParams.has('data')) {
          setToken();
          
        }
        else {
          redirectAuth();
        }
      }
      else {
        if ((Date.now() - parseInt(localStorage.getItem('date'))) > parseInt(localStorage.getItem('expiry')) * 1000) {
          localStorage.removeItem('token');
          localStorage.removeItem('date');
          localStorage.removeItem('expiry');
          redirectAuth();
        }
        else {
          console.log('token still valid');
        }
      }
  
      console.log(localStorage.getItem('token'));
      // console.log(localStorage.getItem('date'));
      // console.log(localStorage.getItem('expiry'));
    }
  }
});
