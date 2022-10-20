<template>
  <div class="q-pa-md">
    <div class="text-h5 q-mb-md">License SKUs</div>
    <q-table
      :rows="o365.skus"
      row-key="id"
      :columns="skuColumns"
      :rows-per-page-options="[10, 20, 30]"
    />
  </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue'
// import { storeToRefs } from 'pinia';
import { useO365Store } from '../stores/o365-store';


const skuColumns = [
  {
    name: 'ID',
    required: true,
    label: 'ID',
    align: 'left',
    field: 'skuId',
    sortable: true
  },
  {
    name: 'Sku Part Number',
    required: true,
    label: 'Sku Part Number',
    align: 'left',
    field: 'skuPartNumber',
    sortable: true
  },
  {
    name: 'Consumed Units',
    required: true,
    label: 'Consumed Units',
    align: 'left',
    field: 'consumedUnits',
    sortable: true
  },
  {
    name: 'Enabled Units',
    required: true,
    label: 'Enabled Units',
    align: 'left',
    field: rows => rows.prepaidUnits.enabled,
    sortable: true
  },
]

export default defineComponent({
  setup() {

    const o365 = useO365Store();
    o365.getGraphSkus();
    o365.auth();
    return {
      // Option 1: return the store directly and couple it in the template
      o365,
      skuColumns

    };
  },
})
</script>
