<template>
  <div class="q-pa-md q-pa-md row items-start q-gutter-md">
    <q-card class="my-card fixed-center">
      <q-card-section>
        <q-input v-model="email" filled type="email" hint="Email" />
        <q-input v-model="password" filled :type="isPwd ? 'password' : 'text'" hint="Password" />
        <br/>
        <div class="row">
            <q-btn class="col-12" color="primary" label="Login" />
        </div>
      </q-card-section>
    </q-card>
  </div>
</template>
<script>
export default {
  data () {
    return {
      password: '',
      isPwd: true,
      email: '',
    }
  }
}
</script>