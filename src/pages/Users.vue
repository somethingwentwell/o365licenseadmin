<template>
  <div class="q-pa-md">
    <div class="row">
      <div class="col-8">
        <q-input v-model="search" filled type="search" hint="Search">
          <template v-slot:append>
            <q-icon name="search" />
          </template>
        </q-input>
      </div>
      <div class="col">
        <q-btn
          :loading="loading1"
          :percentage="percentage1"
          color="primary"
          @click="startUploading(1)"
          style="width: 150px"
          class="full-width"
          size="21px"
        >
          Upload CSV
          <template v-slot:loading>
            <q-spinner-gears class="on-left" />
            Uploading...
          </template>
        </q-btn>
      </div>
      <div class="col">
        <q-btn
          :loading="loading2"
          :percentage="percentage2"
          color="primary"
          @click="startUploading(2)"
          style="width: 150px"
          class="full-width"
          size="21px"
        >
          Download CSV
          <template v-slot:loading>
            <q-spinner-gears class="on-left" />
            Downloading...
          </template>
        </q-btn>
      </div>
    </div>
    <br>
    <q-table
      title="Users"
      :data="data"
      :columns="columns"
      row-key="name"
      :pagination="initialPagination"
      selection="multiple"
      :selected.sync="selected"
    >
      <template v-slot:header-selection="scope">
        Enabled<q-toggle v-model="scope.selected" />
      </template>

      <template v-slot:body-selection="scope">
        Enabled<q-toggle v-model="scope.selected" />
      </template>
    </q-table>
    <q-btn-group spread>
        <q-btn color="primary" label="Apply" />
        <q-btn color="black" label="Reset" />
    </q-btn-group>
  </div>
</template>
<script>
export default {
  data () {
    return {
      loading1: false,
      percentage1: 0,
      loading2: false,
      percentage2: 0,
      selected: [],
      columns: [
        { name: 'assignedLicenses', label: 'assignedLicenses', field: 'assignedLicenses', sortable: true, sort: (a, b) => parseInt(a, 10) - parseInt(b, 10) },
        { name: 'displayName', align: 'center', label: 'displayName', field: 'displayName', sortable: true },
        { name: 'mail', label: 'mail', field: 'mail', sortable: true },
        { name: 'mailNickname', label: 'mailNickname', field: 'mailNickname' },
        { name: 'otherMails', label: 'otherMails', field: 'otherMails' },
        { name: 'proxyAddresses', label: 'proxyAddresses', field: 'proxyAddresses' },
        { name: 'userPrincipalName', label: 'userPrincipalName', field: 'userPrincipalName', sortable: true, sort: (a, b) => parseInt(a, 10) - parseInt(b, 10) }
        
      ],
      data: [
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        },
        {
          assignedLicenses: 'service-value',
          displayName: 'Demo User',
          mail: 'demouser@gmail.com',
          mailNickname: 'contoso1_gmail.com#EXT#',
          otherMails: 'contoso1@gmail.com',
          proxyAddresses: 'SMTP:contoso1@gmail.com',
          userPrincipalName: 'contoso1_gmail.com#EXT#@microsoft.onmicrosoft.com'
        }
      ]
    }
  },
  methods: {
    startUploading (id) {
      this[`loading${id}`] = true
      this[`percentage${id}`] = 0
      this[`interval${id}`] = setInterval(() => {
        this[`percentage${id}`] += Math.floor(Math.random() * 8 + 10)
        if (this[`percentage${id}`] >= 100) {
          clearInterval(this[`interval${id}`])
          this[`loading${id}`] = false
        }
      }, 700)
    }
  },

  beforeDestroy () {
    clearInterval(this.interval1)
  }
}
</script>