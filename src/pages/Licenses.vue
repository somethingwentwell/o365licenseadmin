
<template>
  <div class="q-pa-md">
    <q-table
      title="Treats"
      :data="data"
      :columns="columns"
      row-key="name"
      :visible-columns="visibleColumns"
    >
      <template v-slot:top="props">
        <div class="col-2 q-table__title">Licenses</div>

        <q-space />

        <div v-if="$q.screen.gt.xs" class="col">
          <q-toggle v-model="visibleColumns" val="consumedUnits" label="consumedUnits" />
          <q-toggle v-model="visibleColumns" val="enabled" label="enabled" />
          <q-toggle v-model="visibleColumns" val="suspended" label="suspended" />
          <q-toggle v-model="visibleColumns" val="warning" label="warning" />
        </div>
        <q-select
          v-else
          v-model="visibleColumns"
          multiple
          borderless
          dense
          options-dense
          :display-value="$q.lang.table.columns"
          emit-value
          map-options
          :options="columns"
          option-value="name"
          style="min-width: 150px"
        />

        <q-btn
          flat round dense
          :icon="props.inFullscreen ? 'fullscreen_exit' : 'fullscreen'"
          @click="props.toggleFullscreen"
          class="q-ml-md"
        />
      </template>

    </q-table>
  </div>
</template>
<script>
export default {
  data () {
    return {
      visibleColumns: [ 'servicePlanId', 'servicePlanName', 'provisioningStatus', 'appliesTo', 'consumedUnits', 'enabled', 'suspended', 'warning' ],
      columns: [
        {
          align: 'left',
          field: row => row.name,
          format: val => `${val}`,
          sortable: true
        },
        { name: 'servicePlanId', align: 'center', label: 'servicePlanId', field: 'servicePlanId', sortable: true },
        { name: 'servicePlanName', label: 'servicePlanName', field: 'servicePlanName', sortable: true },
        { name: 'provisioningStatus', label: 'provisioningStatus', field: 'provisioningStatus' },
        { name: 'appliesTo', label: 'appliesTo', field: 'appliesTo' },
        { name: 'consumedUnits', label: 'consumedUnits', field: 'consumedUnits' },
        { name: 'enabled', label: 'enabled', field: 'enabled', sortable: true, sort: (a, b) => parseInt(a, 10) - parseInt(b, 10) },
        { name: 'suspended', label: 'suspended', field: 'suspended', sortable: true, sort: (a, b) => parseInt(a, 10) - parseInt(b, 10) },
        { name: 'warning', label: 'warning', field: 'warning', sortable: true, sort: (a, b) => parseInt(a, 10) - parseInt(b, 10) }
      ],
      data: [
        {
          servicePlanId: 'servicePlanId-value',
          servicePlanName: 'servicePlanName-value',
          provisioningStatus: 'provisioningStatus-value',
          appliesTo: 'appliesTo-value',
          consumedUnits: 99,
          enabled: 99,
          suspended: 99,
          warning: 99
        },
        {
          servicePlanId: 'servicePlanId-value',
          servicePlanName: 'servicePlanName-value',
          provisioningStatus: 'provisioningStatus-value',
          appliesTo: 'appliesTo-value',
          consumedUnits: 99,
          enabled: 99,
          suspended: 99,
          warning: 99
        },
        {
          servicePlanId: 'servicePlanId-value',
          servicePlanName: 'servicePlanName-value',
          provisioningStatus: 'provisioningStatus-value',
          appliesTo: 'appliesTo-value',
          consumedUnits: 99,
          enabled: 99,
          suspended: 99,
          warning: 99
        },
        {
          servicePlanId: 'servicePlanId-value',
          servicePlanName: 'servicePlanName-value',
          provisioningStatus: 'provisioningStatus-value',
          appliesTo: 'appliesTo-value',
          consumedUnits: 99,
          enabled: 99,
          suspended: 99,
          warning: 99
        }
      ]
    }
  }
}
</script>