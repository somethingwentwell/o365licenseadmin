<template>
  <div class="q-pa-md">
    <h5>Logs</h5>
    <q-btn-group spread>
        <q-btn color="primary" label="Download Logs" />
    </q-btn-group>
    <q-list bordered padding>
      <q-item>
        <q-item-section>
          <q-item-label overline>OVERVIEW</q-item-label>
          <q-item-label>Automation logs</q-item-label>
          <q-item-label caption>Successfully assigned 5 license(s)</q-item-label>
          <q-item-label caption>Failed assigned 2 licenses. Fail operation ID [74be27de-1e4e-49d9-b579-fe0b331d3642, 74be27de-1e4e-49d9-b579-fe0b331d3642]</q-item-label>
        </q-item-section>

        <q-item-section side top>
          <q-item-label caption>1 day ago</q-item-label>
        </q-item-section>
      </q-item>

      <q-separator spaced />

      <q-item>
        <q-item-section>
          <q-item-label>Manual operation logs</q-item-label>
          <q-item-label caption>Successfully assigned 1 license(s)</q-item-label>
        </q-item-section>

        <q-item-section side top>
          <q-item-label caption>2 days ago</q-item-label>
        </q-item-section>
      </q-item>

      <q-separator spaced inset="item" />

      <q-item>
        <q-item-section>
          <q-item-label>Manual operation logs</q-item-label>
          <q-item-label caption>Successfully disable 1 license(s)</q-item-label>
        </q-item-section>

        <q-item-section side top>
          <q-item-label caption>3 days ago</q-item-label>
        </q-item-section>
      </q-item>

      <q-separator spaced inset="item" />

      <q-item>
        <q-item-section>
          <q-item-label>Automation logs</q-item-label>
          <q-item-label caption>Successfully assigned 2 license(s)</q-item-label>
          <q-item-label caption>Failed assigned 2 licenses. Fail operation ID [74be27de-1e4e-49d9-b579-fe0b331d3642, 74be27de-1e4e-49d9-b579-fe0b331d3642]</q-item-label>
        </q-item-section>

        <q-item-section side top>
          <q-item-label caption>4 days ago</q-item-label>
        </q-item-section>
      </q-item>

      <q-separator spaced inset="item" />

      <q-item>
        <q-item-section>
          <q-item-label>Automation logs</q-item-label>
          <q-item-label caption>Successfully assigned 1 license(s)</q-item-label>
          <q-item-label caption>Failed assigned 1 licenses. Fail operation ID [74be27de-1e4e-49d9-b579-fe0b331d3642]</q-item-label>
        </q-item-section>

        <q-item-section side top>
          <q-item-label caption>5 days ago</q-item-label>
        </q-item-section>
      </q-item>
    </q-list>
  </div>
</template>