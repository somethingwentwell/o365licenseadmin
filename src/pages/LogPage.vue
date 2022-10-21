<template>
  <div class="q-pa-md">
    <div class="text-h5 q-mb-md">Top 50 Logs</div>
    <q-table
      :rows="o365.logs"
      row-key="id"
      :columns="logColumns"
      :rows-per-page-options="[10, 20, 30]"
    >
      <template v-slot:top-right>
        <q-input borderless dense debounce="300" v-model="filter" placeholder="Search">
          <template v-slot:append>
            <q-icon name="search" />
          </template>
        </q-input>
      </template>
    </q-table>
  </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue'
// import { storeToRefs } from 'pinia';
import { useO365Store } from '../stores/o365-store';
import { ref } from 'vue';

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
    name: 'UserId',
    required: true,
    label: 'UserId',
    align: 'left',
    field: 'UserId',
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
    name: 'Request Body',
    required: true,
    label: 'Request Body',
    align: 'left',
    field: 'RequestBody',
    sortable: true
  },
  {
    name: 'Result',
    required: true,
    label: 'Result',
    align: 'left',
    field: row => {
      if (row.Result) {
        return row.Result;
      }
      return '';
    },
    sortable: true
  }
]

export default defineComponent({
  setup() {

    const o365 = useO365Store();
    o365.getLogs();
    return {
      // Option 1: return the store directly and couple it in the template
      o365,
      logColumns,
      filter: ref(''),

    };
  },
})
</script>
