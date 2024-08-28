<template>
  <div id="app">
    <h1>Excel File Import and Export</h1>

    <input type="file" @change="importExcelFile" accept=".xlsx" />
    <button @click="exportExcelFile(myData, 'vertical')">
      Export to Excel
    </button>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import { useExcelJS } from "../composables/useExcelJS";
import axios from "axios";

const { importExcelFile, exportExcelFile } = useExcelJS();

const myData = ref<any[]>([]);

const getData = async () => {
  try {
    const { data, status } = await axios({
      method: "POST",
      url: "http://217.174.224.134:6888/api/v1/gumruk/admin/reports/get-reports",
      data: {
        page: 0,
        limit: 20,
        search: "",
        postId: "",
        employeeId: "",
        dateTo: "2024 - 05 - 01",
        dateFrom: "2024 - 08 - 19",
      },
    });
    if (status) {
      myData.value = data.data.map((item: any) => {
        return {
          fullName: `${item.employeeName} ${item.employeeSurName} ${item.employeeLastName}`,
          job: item.employeeProfession,
          pointCode: item.pointNumber,
          pointAddress: item.pointName,
        };
      });

    }
  } catch (e) {}
};

getData();
</script>

<style>
#app {
  text-align: center;
  font-family: Avenir, Helvetica, Arial, sans-serif;
  color: #0055ff;
  margin-top: 60px;
}
input[type="file"] {
  margin: 20px;
}
button {
  margin: 20px;
}
</style>