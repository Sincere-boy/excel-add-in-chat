<script setup>
import {defineComponent} from "vue";
import HeaderNav from "@/HeaderNav.vue";
import Message from "@/conponents/Message.vue";

const height = window.innerHeight;
const alayoutstyle = {
height: height - 20 + 'px',
padding: 0,
display: 'flex',
}

</script>

<template>
  <div class="common-layout" :style="alayoutstyle">
    <el-container style="height: 100%">
      <el-header class="header-chat">
        <HeaderNav />
      </el-header>
<!--      <el-container style="padding: 0!important; flex-direction: column !important;">-->
      <Message />
<!--      </el-container>-->
    </el-container>
  </div>
<!--  <div class="common-layout">-->
<!--    <el-container>-->
<!--      <el-header>Header</el-header>-->
<!--      <el-main>Main</el-main>-->
<!--    </el-container>-->
<!--  </div>-->
</template>

<style scoped>
.header-chat {
  padding: 0 !important;
  height: 30px;
}
.body {
  padding: 0 !important;
  flex-grow: 1;
}

</style>