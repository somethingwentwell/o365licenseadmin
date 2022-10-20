<template>
  <div class="q-pa-md">
    <div class="text-h5 q-mb-md">Rules</div>
    <div class="q-gutter-md column" style="max-width: 1000px">

      <div class="text-h6">
        Email for Automation Fail Alert
      </div>
      <q-input outlined v-model="o365.config.email" />

      <div class="text-h6">
        Deactivate Period:: {{ o365.config.deactivate_period }} (0 to 90)
      </div>
      <q-slider v-model="o365.config.deactivate_period" :min="0" :max="90"/>


      <div class="text-h6">
        Delete Period:: {{ o365.config.delete_period }} (0 to 90)
        </div>
      <q-slider v-model="o365.config.delete_period" :min="0" :max="90" color="green"/>
    </div>
    <br/>
    <div class="qm-pa-md q-gutter-s">
      <q-btn color="primary" label="Save" @click="updateConfig" />
    </div>
  </div>
</template>
<script lang="ts">
import { defineComponent } from 'vue'
import { useO365Store } from '../stores/o365-store';

  const o365 = useO365Store();
  o365.getConfig();

let updateConfig = () => {
  o365.updateConfig();
}

export default defineComponent({
  setup() {
    o365.auth();
    return {
      o365,
      updateConfig
    };
  },
})
</script>
