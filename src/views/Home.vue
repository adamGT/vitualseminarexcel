<template>
  <div>
    <div>
      
      <div v-if="isDataSaved"><h1 class="seminar-title">{{seminarTitle}}</h1></div>
      <div v-if="isDataSaved"><p class="presenter-name">{{presentersName}}</p></div>
      <v-row v-if="!isDataSaved">
        <v-col
          class="d-flex"
          cols="8"
          md="4"
          sm="4"
        >
          <v-text-field
            class="mt-4"
            v-model="seminarTitle"
            label="Seminar Title"
            dense
            single-line
            outlined
          ></v-text-field>
        </v-col>
        
        <v-col
          class="d-flex"
          cols="8"
          md="4"
          sm="4"
        >
          <v-text-field
            class="mt-4"
            v-model="presentersName"
            label="Presenter's Name"
            dense
            single-line
            outlined
          ></v-text-field>
        </v-col>
        <v-col
          class="d-flex"
          cols="8"
          md="4"
          sm="4"
          >
          <v-btn
            class="mt-4"
            rounded
            color="primary"
            @click="saveData"
            dark
          >
            Save Data
          </v-btn>
        </v-col>
      </v-row>
      <v-row>
        
        <v-col
          class="d-flex"
          cols="8"
          md="3"
          sm="3"
        >
          <div>
            <span class="btn btn-primary btn-file">
                Upload File <input type="file" @change="onChange">
            </span>
          </div>
        </v-col>
        
        
      </v-row>
      
    </div>
    <div
      v-if="file"
    >
      <h4> Choose a sheet</h4>
      <v-row>
        <v-col
          class="d-flex"
          cols="8"
          md="3"
          sm="3"
        >
          <v-select
            :items="sheets"
            v-model="selectedSheet"
            label="Choose a sheet"
            outlined
          ></v-select>
        </v-col>
        
        <v-col
          class="d-flex"
          cols="8"
          md="2"
          sm="3"
        >
          <v-btn 
            class="ma-2"
            @click="displayTable"
            color="info"
            :disabled="selectedSheet === null"
            > Display Data</v-btn>
        </v-col>
        <v-col
          class="d-flex"
          cols="8"
          md="2"
          sm="3"
        >
          <v-btn 
            class="ma-2"
            @click="sendemail"
            color="success"
            :disabled="selected.length === 0"
            > Send Email</v-btn>
        </v-col>
        <v-col
          class="d-flex"
          cols="8"
          md="2"
          sm="3"
        >
          <v-btn 
            @click="exportExcel"
            class="ma-2"
            color="success"
            :disabled="!isDataChanged"
            > Download Data</v-btn>
        </v-col>
      </v-row>
    </div>
    
    <v-data-table
      v-model="selected"
      v-if="displaytable"
      :headers="headers"
      :items="collection"
      :items-per-page="5"
      class="elevation-1"
      :search="search"
      item-key="Name"
      :custom-filter="filterOnlyCapsText"
      show-select
      :single-select="true"
    >
      <template v-slot:top>
        <v-text-field
          v-model="search"
          label="Search (UPPER CASE ONLY)"
          class="mx-4"
        ></v-text-field>
      </template>

      <template v-slot:[`item.Paid`]="{ item }">
        <v-chip
          :color="getColor(item.Paid)"
          dark
        >
          {{ item.Paid }}
        </v-chip>
      </template>
    </v-data-table>
  </div>
</template>

<script>
import * as XLSX from 'xlsx'

export default {
  components: {
  },
  data() {
    return {
      search: '',
      file: null,
      selectedSheet: null,
      selected: [],
      sheetName: null,
      sheets: [],
      seminarTitle: null,
      presentersName: null,
      isDataSaved: false,
      isDataChanged: false,
      collection: [],
      displaytable: false,
      headers: [
        { text: 'Name', value: 'Name' },
        { text: 'Email', value: 'Email' },
        { text: 'Number', value: 'Number' },
        { text: 'Date', value: 'Date' },
        { text: 'Link', value: 'Link' },
        { text: 'Paid', value: 'Paid' },
      ],
      mail: "Dear ,Thank you for participating in our virtual seminar on Medical Error and Negligence by Dr. Abnet Ayalew (MD).We have received your payment for the certificate. Please find the attached file of your certificate below. and you can share your certificate on LinkedIn and other social media. You can also get the recorded session on our YOUTUBE (Subscribe to get previous and future sessions) by clicking the link below. YouTube channel link https://www.youtube.com/channel/UC2v9R3kHD4Auor0NkyQYHZw Link to your certificate: Best regards, Blue Health team"
    };
  },
  methods: {
    displayData(){
      console.log(this.sheets)
    },
    onChange(event) {
      this.file = event.target.files ? event.target.files[0] : null;
      if (this.file) {
        const reader = new FileReader();

        reader.onload = (e) => {
          /* Parse data */
          const bstr = e.target.result;
          const wb = XLSX.read(bstr, { type: 'binary' });
          /* Get first worksheet */
          const sheetsname = wb.SheetNames;
          const wsname = wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          /* Convert array of arrays */ 
          // const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
          const data = XLSX.utils.sheet_to_json(ws);
          this.setData(data, sheetsname)
        }

        reader.readAsBinaryString(this.file);
      }
      
    },
    setData(data, sheets){
      this.collection = data;
      this.sheets = sheets;
    },
    onButton(){
      console.log(this.collection)
    },
    addSheet() {
      this.sheets.push({ name: this.sheetName, data: [...this.collection] });
      this.sheetName = null;
    },
    displayTable() {
      this.header=this.collection[0]
      this.displaytable = true;
    },
    filterOnlyCapsText (value, search) {
      return value != null &&
        search != null &&
        typeof value === 'string' &&
        value.toString().toLocaleUpperCase().indexOf(search) !== -1
    },
    sendemail(){
      for(let i = 0; i < this.selected.length; i++){
        window.open(`mailto:${this.selected[i].Email}?subject=1.5 CEU CPD certificate for ${this.seminarTitle} Virtual Seminar&body=Dear ${this.selected[i].Name} , %0D%0DThank you for participating in our virtual seminar on the ${this.seminarTitle} by ${this.presentersName}. We have received your payment for the certificate. %0DPlease find the attached file of your certificate below. and you can share your certificate on LinkedIn and other social media. You can also get the recorded session on our YOUTUBE (Subscribe to get previous and future sessions) by clicking the link below. %0D%0DYouTube channel link %0Dhttps://www.youtube.com/channel/UC2v9R3kHD4Auor0NkyQYHZw %0D%0DLink to your certificate: ${this.selected[i].Link}   %0D%0DBest regards, %0DBlue Health team`);
        this.collection[this.collection.indexOf(this.selected[i])].Paid = "yes"
        this.isDataChanged = true
      }
    },
    saveData(){
      this.isDataSaved=true;
    },
    getColor(paid){
      if(paid === "yes"){
        return "green";
      }else {
        return "orange";
      }
    },
    exportExcel(){
      this.isDataChanged = false
      const worksheet = XLSX.utils.json_to_sheet(this.collection);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "CERTIFICATE lIST");
      XLSX.writeFile(workbook, "CERTIFICATES.xlsx");
    }
  }
};
</script>

<style lang="scss">
.seminar-title {
    font-size: 4rem;
  }
.presenter-name{
  font-size: 3rem;
}
.custom-file-upload {
    border: 1px solid #ccc;
    display: inline-block;
    padding: 6px 12px;
    cursor: pointer;
}
.btn-file {
  position: relative;
  overflow: hidden;
}
.btn-file input[type=file] {
  position: absolute;
  top: 0;
  right: 0;
  min-width: 100%;
  min-height: 100%;
  font-size: 100px;
  text-align: right;
  filter: alpha(opacity=0);
  opacity: 0;
  outline: none;
  background: white;
  cursor: inherit;
  display: block;
}
</style>