<template>
  <div id="app">
    <Demo />
    <Demo1 />
  </div>
</template>

<script>
import Demo from './components/Demo.vue'
import Demo1 from './components/Demo1.vue'

export default {
  name: 'app',
  components: {
    Demo,
    Demo1
  }
}
</script>

