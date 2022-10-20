<template>
  <div class="q-pa-md">
    <div class="text-h5 q-mb-md">Top 50 Logs</div>
    <q-table
      :rows="o365.logs"
      row-key="id"
      :columns="logColumns"
      :rows-per-page-options="[10, 20, 30]"
    />
  </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue'
// import { storeToRefs } from 'pinia';
import { useO365Store } from '../stores/o365-store';


const logColumns = [
  // {
  //   name: 'ID',
  //   required: true,
  //   label: 'ID',
  //   align: 'left',
  //   field: 'RowKey',
  //   sortable: true
  // },
  {
    name: 'Action',
    required: true,
    label: 'Action',
    align: 'left',
    field: 'PartitionKey',
    sortable: true
  },
  {
    name: 'Timestamp',
    required: true,
    label: 'Timestamp',
    align: 'left',
    field: 'Timestamp',
    sortable: true
  },
  {
    name: 'Success',
    required: true,
    label: 'Success',
    align: 'left',
    field: 'success',
    sortable: true
  },
  {
    name: 'Error Message',
    required: true,
    label: 'Error Message',
    align: 'left',
    field: row => {
      let obj = JSON.parse(row.Result);
      if (obj['error']) {
        return obj['error']['message'];
      }
      return '';
    },
    sortable: true
  },
  {
    name: 'Request Body',
    required: true,
    label: 'Request Body',
    align: 'left',
    field: 'RequestBody',
    sortable: true
  }
]

export default defineComponent({
  setup() {

    const o365 = useO365Store();
    o365.auth();
    o365.getLogs();
    return {
      // Option 1: return the store directly and couple it in the template
      o365,
      logColumns

    };
  },
})
</script>
