<template>
  <div class="q-pa-md">
    <div class="text-h5 q-mb-md">Users - Current: {{ o365.currentUser.id }} </div>
    <q-table
      :rows="o365.customUsers"
      row-key="id"
      :columns="userColumns"
      :rows-per-page-options="[10, 20, 30]"
      :filter="filter"
    >
      <template v-slot:top-left>
      <q-uploader
        style="max-width: 300px"
        :url="blobUrl"
        method="PUT"
        :headers="[{name: 'x-ms-blob-type', value: 'BlockBlob'}]"
        label="Upload csv"
        accept=".csv"
      />
        <q-btn
          style="width: 300px"
          color="primary"
          icon-right="archive"
          label="Export to csv"
          no-caps
          @click="exportCsv"
        />
      </template>
      <template v-slot:top-right>
        <q-input borderless dense debounce="300" v-model="filter" placeholder="Search">
          <template v-slot:append>
            <q-icon name="search" />
          </template>
        </q-input>
      </template>
      <template v-slot:header="props">
        <q-tr :props="props">
          <q-th auto-width />
          <q-th
            v-for="col in props.cols"
            :key="col.name"
            :props="props"
          >
            {{ col.label }}
          </q-th>
        </q-tr>
      </template>

      <template v-slot:body="props">
        <q-tr :props="props">
          <q-td auto-width>
            <q-btn size="sm" color="accent" round dense @click="setCurrentUser(props.row.id); props.expand = !props.expand" :icon="props.expand ? 'remove' : 'add'" />
          </q-td>
          <q-td
            v-for="col in props.cols"
            :key="col.name"
            :props="props"
          >
            {{ col.value }}
          </q-td>
        </q-tr>
        <q-tr v-show="props.expand" :props="props">
          <q-td colspan="100%">
            <div class="text-left">
              <div class="q-pa-sm">
                <div class="q-pa-md">
                  <q-option-group
                    :options="options"
                    type="checkbox"
                    v-model="group"
                  />
                </div>
                <div class="q-px-sm">
                  Selected SKUs: <strong>{{ group }}</strong>
                </div>
              </div>
              <div class="q-pa-md">
                <q-btn color="primary" label="Add License(s)" @click="addLicense()" />
              </div>
              <q-table
                :rows="props.row.assignedLicenses"
                row-key="servicePlanId"
                :columns="licenseColumns"
              >
              <template v-slot:header="props">
                <q-tr :props="props">
                  <q-th auto-width />
                  <q-th
                    v-for="col in props.cols"
                    :key="col.name"
                    :props="props"
                  >
                    {{ col.label }}
                  </q-th>
                </q-tr>
              </template>
              <template v-slot:body="props">
                <q-tr :props="props">
                  <q-td auto-width>
                    <q-btn size="sm" label="remove" color="red" @click="removeLicense(props.row)" />
                  </q-td>
                  <q-td
                    v-for="col in props.cols"
                    :key="col.name"
                    :props="props"
                  >
                    {{ col.value }}
                  </q-td>
                </q-tr>
              </template>
              </q-table>
            </div>
          </q-td>
        </q-tr>
      </template>
    </q-table>
  </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue'
// import { storeToRefs } from 'pinia';
import { useO365Store } from '../stores/o365-store';
import { ref } from 'vue';
import { exportFile, useQuasar } from 'quasar'
const $q = useQuasar();
const o365 = useO365Store();
o365.getGraphCustomUsers();
o365.getGraphSkus();

let options = ref([]);
let group = ref([]);


const userColumns = [
  {
    name: 'ID',
    required: true,
    label: 'ID',
    align: 'left',
    field: 'id',
    sortable: true
  },
  {
    name: 'Display Name',
    required: true,
    label: 'Display Name',
    align: 'left',
    field: 'displayName',
    sortable: true
  },
  {
    name: 'Mail',
    required: true,
    label: 'Mail',
    align: 'left',
    field: 'mail',
    sortable: true
  },
  {
    name: 'Last Sign In',
    required: true,
    label: 'Last Sign In',
    align: 'left',
    field: row => {
      return row.signInActivity ? new Date(row.signInActivity.lastSignInDateTime).toLocaleString() : '';
    },
    sortable: true
  },
  {
    name: 'Job Title',
    required: true,
    label: 'Job Title',
    align: 'left',
    field: 'jobTitle',
    sortable: true
  },
  {
    name: 'Office Location',
    required: true,
    label: 'Office Location',
    align: 'left',
    field: 'officeLocation',
    sortable: true
  },
]

const licenseColumns = [
  {
    name: 'SKU ID',
    required: true,
    label: 'SKU ID',
    align: 'left',
    field: 'skuId',
    sortable: true
  },
  {
    name: 'Sku Part Number',
    required: true,
    label: 'Sku Part Number',
    align: 'left',
    field: row => {
      let skuPartNumber = '';
      o365.skus.forEach((sku) => {
        if (sku.skuId === row.skuId) {
          skuPartNumber = sku.skuPartNumber;
        }
      });
      return skuPartNumber;
    },
    sortable: true
  }
]

let setCurrentUser = (userId) => {
  o365.currentUser.id = userId;
  let skuOptions = [];
    o365.getSkus.forEach(element => {
    skuOptions.push({ label: element.skuPartNumber, value: element.skuId });
  });
  options.value = skuOptions;
  }

let removeLicense = (row) => {
  o365.removeUserLicense(o365.currentUser.id, row.skuId);
}

let addLicense = () => {
  let formattedGroup = [] as string[];
  group.value.forEach((element) => {
    formattedGroup.push(element);
  });
  o365.addUserLicense(o365.currentUser.id, formattedGroup);
}

let exportCsv = () => {

  const items = o365.customUsers;
  const replacer = (key, value) => value === null ? '' : value // specify how you want to handle null values here
  const header = Object.keys(items[0])
  header.join(',');

  const csv = [
    header.join(','), // header row first
    ...items.map(row => header.map(fieldName => {
      if (fieldName === 'assignedLicenses') {
        //join the array of objects into a string with comma
        return `${row[fieldName].map((license) => license.skuId).join(';')}`;
      } else if (fieldName === 'signInActivity') {
        if (row[fieldName]) {
          return `"${new Date(row[fieldName].lastSignInDateTime)}"`;
        } else {
          return '';
        }
      } else {
        return JSON.stringify(row[fieldName], replacer)
      }
    }).join(','))
  ].join('\r\n')

  const status = exportFile(
    'user-export.csv',
    csv,
    'text/csv'
  )

  if (status !== true) {
    $q.notify({
      message: 'Browser denied file download...',
      color: 'negative',
      icon: 'warning'
    })
  }
}

export default defineComponent({
  setup() {
    o365.auth();
    return {
      o365,
      userColumns,
      licenseColumns,
      setCurrentUser,
      removeLicense,
      addLicense,
      group,
      options,
      exportCsv,
      blobUrl: `https://${o365.storage.accountName}.blob.core.windows.net/users/users.csv?${localStorage.getItem('sas_token')}`,
      filter: ref(''),

    };
  },
})
</script>
