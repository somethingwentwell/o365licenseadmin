<template>
  <div>
  </div>
</template>

<script lang="ts">

import { defineComponent } from 'vue';
import { useO365Store } from '../stores/o365-store';
const o365 = useO365Store();

export default defineComponent({
  setup() {
    o365.auth();
    return {
      o365
    };
  },
})
</script>
