
<template>
  <q-layout view="hHh lpR fFf">

    <q-header elevated class="bg-primary text-white">
      <q-toolbar>
        <q-btn dense flat round icon="menu" @click="left = !left" />

        <q-toolbar-title>
          <q-avatar>
            <img src="icons/favicon-96x96.png">
          </q-avatar>
          Kai Shing O365 Licenses Management
        </q-toolbar-title>
      </q-toolbar>
    </q-header>

    <q-drawer show-if-above v-model="left" side="left" bordered>
      <!-- drawer content -->
        <q-scroll-area class="fit">
          <q-list padding class="menu-list">
            <q-item clickable v-ripple @click="$router.replace('/users')">
            
              <q-item-section avatar>
                <q-icon name="people" />
              </q-item-section>

              <q-item-section>
                Users
              </q-item-section>
            </q-item>

            <q-item clickable v-ripple @click="$router.replace('/skus')">
              <q-item-section avatar>
                <q-icon name="subscriptions" />
              </q-item-section>

              <q-item-section>
                Licenses
              </q-item-section>
            </q-item>

            <q-item clickable v-ripple @click="$router.replace('/rules')">
              <q-item-section avatar>
                <q-icon name="rule" />
              </q-item-section>

              <q-item-section>
                Automation Rules
              </q-item-section>
            </q-item>

            <q-item clickable v-ripple @click="$router.replace('/logs')">
              <q-item-section avatar>
                <q-icon name="summarize" />
              </q-item-section>

              <q-item-section>
                Logs
              </q-item-section>
            </q-item>

          </q-list>
        </q-scroll-area>
    </q-drawer>

    <q-page-container>
      <router-view />
    </q-page-container>

  </q-layout>
</template>

<script>
export default {
  data () {
    return {
      left: false
    }
  }
}
</script>