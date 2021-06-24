<template>
  <div class="q-pa-md">
    <h5>Rules</h5>
    <div class="q-gutter-md column" style="max-width: 1000px">

      <q-field standout square hint="CSV Upload and Read Path" :dense="dense">
        <template v-slot:control>
          <div class="self-center full-width no-outline" tabindex="0">{{text}}</div>
        </template>
      </q-field>

      <q-field standout square hint="Email for Automation Fail Alert" :dense="dense">
        <template v-slot:control>
          <div class="self-center full-width no-outline" tabindex="0">{{text2}}</div>
        </template>
      </q-field>

      <q-field
        filled
        :hint="`${ slider } days to disable inactive O365 license`"
        :value="slider"
        @input="val => (val === null && (slider = 60))"
        clearable
      >
        <template v-slot:control>
          <q-slider
            :value="slider"
            @change="val => { slider = val }"
            :min="0"
            :max="100"
            label
            label-always
            class="q-mt-lg"
          />
        </template>
      </q-field>
      <q-field
        filled
        :hint="`${ slider2 } days to delete disable O365 license`"
        :value="slider2"
        @input="val => (val === null && (slider2 = 60))"
        clearable
      >
        <template v-slot:control>
          <q-slider
            :value="slider2"
            @change="val => { slider2 = val }"
            :min="0"
            :max="100"
            label
            label-always
            class="q-mt-lg"
          />
        </template>
      </q-field>
    </div>
  </div>
</template>
<script>
export default {
  data () {
    return {
      slider: 60,
      slider2: 60,
      text: 'Field content',
      text2: 'Field content',
      dense: false
    }
  }
}
</script>