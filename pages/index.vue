<template>
  <v-row justify="center" align="center">
    <v-col style="padding: 10px;">
      <v-card>
        <v-card-title class="headline">
          <div style="padding-right: 4rem;">
            <h1>Time Sheet</h1>
          </div>
          <br />
          <br />
          <div>
            <span><b>Name:</b> {{ fullName }}</span><br />
            <span><b>Position:</b> {{ position }}</span>
          </div>
        </v-card-title>
        <v-card style="padding: 0rem 4rem;">
          <div>
          <v-treeview rounded hoverable  :items="items">
            <template v-slot:prepend="{ item }">
              <div>
                <div v-if="!!item.children">
                  <v-spacer />
                  <v-btn color="info" @click="addNewField(item.id)">
                    +
                  </v-btn>
                </div>
                <div v-else style="padding: 10px 0px;">
                  <p>{{ item.index + 1 }}.</p>
                  <v-row>
                    <v-col cols="12" sm="6" md="3">
                      <v-select hide-details :items="projectNameItems" label="Project Name"
                        v-model="item.data[2]" outlined></v-select>
                    </v-col>
                    <v-col cols="12" sm="6" md="3">
                      <v-select hide-details :items="projectStageItems" label="Project  Stage" outlined
                        v-model="item.data[3]"></v-select>
                    </v-col>
                    <v-col cols="12" sm="6" md="3">
                      <v-select hide-details :items="statusItems" label="Status" outlined
                        v-model="item.data[4]">
                      </v-select>
                    </v-col>
                    <v-col cols="12" sm="6" md="3">
                      <v-select hide-details :items="deliveredToItems" label="Delivered to" outlined
                        v-model="item.data[5]"></v-select>
                    </v-col>
                    <v-col cols="12" sm="6" md="4">
                      <v-text-field hide-details label="Notes" outlined v-model="item.data[6]">
                      </v-text-field>
                    </v-col>
                    <v-col cols="12" sm="6" md="8">
                      <v-text-field hide-details label="Task description" outlined v-model="item.data[7]">
                      </v-text-field>
                    </v-col>
                  </v-row>
                  <br>
                  <hr/>
                </div>
              </div>
            </template>
          </v-treeview>
        </div>
        </v-card>
       
        <v-card-actions>
          <v-spacer />
          <v-btn color="warning" nuxt to="/inspire">
            Generate to EXCEL
          </v-btn>
        </v-card-actions>
      </v-card>
    </v-col>
  </v-row>
</template>

<script>
  export default {
    name: 'IndexPage',
    data: () => ({
      fullName: 'FristName LastName',
      position: 'Programmer',
      projectNameItems: ['Commission System', 'Cost Sheet Management System', 'Supplier Portal System',
        'Excel Converter', 'Time Attendance System', 'All Dela', 'River-Leaf'
      ],
      projectStageItems: ['Meeting (project)', 'Project Initiation', 'Project Planning', 'Requirement', 'Design',
        'Programming', 'Testing', 'Operation Readiness', 'User Acceptance Test', 'Project Closing',
      ],
      statusItems: ['In Progress', 'Complete'],
      deliveredToItems: ['K.Naris', 'K.Ladawan', 'K.Kriangkrai', 'K.Weerapat', 'K.Surachect', 'K.Kitti',
        'K.Benjaporn', 'K.Suphansa', 'K.Kwanthip', 'K.Thanyathip', 'K. Jaroon'
      ],
      items: [{
          id: 0,
          name: 'Hours for Project :',
          children: [{
            index: 0,
            data: ['1', '']
          }],
        },
        {
          id: 1,
          name: 'Hours for Non Project :',
          children: [{
            index: 0,
            data: ['1', '']
          }]
        },
        {
          id: 2,
          name: 'Hours for Training/Seminar :',
          children: [{
            index: 0,
            data: ['','']
          }]
        },
        {
          id: 3,
          name: 'Hours for Leave :',
          children: [{
            index: 0,
            data: ['','']
          }],
        },
      ],
    }),

    methods:{
      addNewField(id){
        let index = this.items[id].children.length + 1
        let indexText = id === 0 || id === 1 ? index.toString() : ''

        this.items[id].children.push({
          index: index - 1 - id,
          data: [indexText ,'']
        })
      }
    }
  }

</script>
