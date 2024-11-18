<template>
  <el-select
    ref="selectInput"
    v-model="tags"
    multiple
    :collapse-tags="collapseTag"
    clearable
    filterable
    allow-create
    default-first-option
    :placeholder="placeholder"
  >
    <el-option v-for="item in tags" :key="item" :label="item" :value="item"> </el-option>
  </el-select>
</template>

<script>
export default {
  props: {
    value: {
      type: String,
      default: '',
    },
    collapseTag: {
      type: Boolean,
      default: true,
    },
    placeholder: {
      type: String,
      default: '请输入',
    },
  },
  data() {
    return {
      tags: this.value.split(',').filter((tag) => tag.trim() !== ''),
    }
  },
  watch: {
    value: {
      handler(newVal) {
        this.tags = newVal.split(',').filter((tag) => tag.trim() !== '')
      },
      immediate: true,
    },
    tags: {
      handler(newVal) {
        this.$emit('input', newVal.join(','))
      },
      deep: true,
    },
  },
  methods: {},
}
</script>
