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
            <v-row>
              <v-col>
                <v-text-field hide-details label="Name: " outlined v-model="fullName">
                </v-text-field>
              </v-col>
              <v-col>
                <v-text-field hide-details label="Position:" outlined v-model="position">
                </v-text-field>
              </v-col>
              <v-col cols="12" sm="6" md="4">
                <v-menu outlined v-model="menu2" :close-on-content-click="false" :nudge-right="40"
                  transition="scale-transition" offset-y min-width="auto">
                  <template v-slot:activator="{ on, attrs }">
                    <v-text-field outlined v-model="date" label="Picker without buttons" prepend-inner-icon="mdi-calendar" readonly
                      v-bind="attrs" v-on="on"></v-text-field>
                  </template>
                  <v-date-picker v-model="date" @input="menu2 = false"></v-date-picker>
                </v-menu>
              </v-col>
            </v-row>
          </div>
        </v-card-title>
        <v-card style="padding:0rem 2rem 1rem 2rem;">
          <div>
            <v-treeview rounded hoverable :items="items">
              <template v-slot:prepend="{ item }">
                <div>
                  <div v-if="!!item.children">
                    <v-spacer />
                    <v-btn color="info" @click="addNewField(item.id)" :disabled="item.id !== 0">
                      +
                    </v-btn>
                  </div>
                  <div v-else style="padding: 10px 0px;">
                    <div v-if="item.parentId === 0">
                      <p>{{ item.index + 1 }}.</p>
                    <v-row>
                      <v-col cols="12" sm="6" md="3">
                        <v-select hide-details :items="projectNameItems" label="Project Name" v-model="item.data[2]"
                          outlined></v-select>
                      </v-col>
                      <v-col cols="12" sm="6" md="3">
                        <v-select hide-details :items="projectStageItems" label="Project  Stage" outlined
                          v-model="item.data[3]"></v-select>
                      </v-col>
                      <v-col cols="12" sm="6" md="3">
                        <v-select hide-details :items="statusItems" label="Status" outlined v-model="item.data[4]">
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
                      <v-col cols="12" sm="6" md="3">
                        <v-text-field type="number" hide-details label="Hour" outlined v-model="item.data[8]">
                        </v-text-field>
                      </v-col>
                      <v-col cols="12" sm="12" md="12">
                        <v-text-field hide-details label="Task description" outlined v-model="item.data[7]">
                        </v-text-field>
                      </v-col>
                    </v-row>
                    <br>
                    <hr />
                    </div>
                    
                   
                  </div>
                </div>
              </template>
            </v-treeview>
          </div>
        </v-card>

        <v-card-actions outlined>
          <v-spacer />
          <v-btn color="warning" @click="generateData()">
            Generate to EXCEL
          </v-btn>
        </v-card-actions>
      </v-card>
    </v-col>
  </v-row>
</template>

<script>
  import * as Excel from 'exceljs';
  import {
    saveAs
  } from "file-saver";

  export default {
    name: 'IndexPage',
    data: () => ({
      fullName: 'FristName LastName',
      position: 'Programmer',
      date: (new Date(Date.now() - (new Date()).getTimezoneOffset() * 60000)).toISOString().substr(0, 10),
      menu: false,
      menu2: false,
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
            parentId: 0,
            index: 0,
            data: ['1', '']
          }],
        },
        {
          id: 1,
          name: 'Hours for Non Project :',
          children: [{
            parentId: 1,
            index: 0,
            data: ['1', '']
          }]
        },
        {
          id: 2,
          name: 'Hours for Non Project :',
          children: [{
            parentId: 2,
            index: 0,
            data: ['1', '']
          }]
        },
        {
          id: 3,
          name: 'Hours for Incident/Bug Fix :',
          children: [{
            parentId: 3,
            index: 0,
            data: ['', '']
          }]
        },
        {
          id: 4,
          name: 'Hours for Leave :',
          children: [{
            parentId: 4,
            index: 0,
            data: ['', '']
          }],
        },
      ],
    }),

    methods: {
      addNewField(id) {
        let index = this.items[id].children.length + 1
        let indexText = id === 0 || id === 1 ? index.toString() : ''

        this.items[id].children.push({
          parentId: id,
          index: index - 1 - id,
          data: [indexText, '']
        })
      },
      async generateData() {
        let projectItems = []
        let nonProjectItems = []
        let incidentItems = []
        let traingingItems = []
        let leaveItems = []
        let totalHoursforProject = 0
        let totalHoursforNonProject = 0
        let totalHoursforIncident = 0
        let totalHoursforTraining = 0
        let totalHoursforLeave = 0
        let totalHours = 0

        this.items[0].children.forEach((element, index) => {
          let hour = element.data[element.data.length - 1]
          element.data[element.data.length] = hour
          totalHoursforProject += Number(hour)

          for (let i = 0; i < element.data.length; i++) {
            element.data[i] = !!element.data[i] ? element.data[i] : "";
          }
          projectItems.push(element)
        });

        this.items[1].children.forEach((element, index) => {
          let hour = element.data[element.data.length - 1]
          element.data[element.data.length] = hour
          totalHoursforNonProject += Number(hour)

          for (let i = 0; i < element.data.length; i++) {
            element.data[i] = !!element.data[i] ? element.data[i] : "";
          }
          nonProjectItems.push(element)
        });

        this.items[2].children.forEach((element, index) => {
          let hour = element.data[element.data.length - 1]
          element.data[element.data.length] = hour
          totalHoursforIncident += Number(hour)

          for (let i = 0; i < element.data.length; i++) {
            element.data[i] = !!element.data[i] ? element.data[i] : "";
          }
          incidentItems.push(element)
        });

        this.items[3].children.forEach((element, index) => {
          let hour = element.data[element.data.length - 1]
          element.data[element.data.length] = hour
          totalHoursforTraining += Number(hour)

          for (let i = 0; i < element.data.length; i++) {
            element.data[i] = !!element.data[i] ? element.data[i] : "";
          }
          traingingItems.push(element)
        });

        this.items[4].children.forEach((element, index) => {
          let hour = element.data[element.data.length - 1]
          element.data[element.data.length - 2] = hour
          totalHoursforLeave += Number(hour)

          for (let i = 0; i < element.data.length; i++) {
            element.data[i] = !!element.data[i] ? element.data[i] : "";
          }
          leaveItems.push(element)
        });


        totalHours = totalHoursforProject + totalHoursforNonProject + totalHoursforIncident + totalHoursforTraining + totalHoursforLeave


        const data = {
          date: this.date,
          fullName: this.fullName,
          position: this.position,
          timeIn: '9:30 AM',
          timeOut: '6:30 PM',
          totalHoursforProject: totalHoursforProject.toString(),
          totalHoursforProjectText: parseFloat(totalHoursforProject).toFixed(2),
          totalHoursforNonProject: totalHoursforNonProject.toString(),
          totalHoursforNonProjectText: parseFloat(totalHoursforNonProject).toFixed(2),
          totalHoursforIncident: totalHoursforIncident.toString(),
          totalHoursforIncidentText: parseFloat(totalHoursforIncident).toFixed(2),
          totalHoursforTraining: totalHoursforTraining.toString(),
          totalHoursforTrainingText: parseFloat(totalHoursforTraining).toFixed(2),
          totalHoursforLeave: totalHoursforLeave.toString(),
          totalHoursforLeaveText: parseFloat(totalHoursforLeave).toFixed(2),
          totalHours: totalHours.toString(),
          totalHoursText: parseFloat(totalHours).toFixed(2),
          projectItems: projectItems,
          nonProjectItems: nonProjectItems,
          incidentItems: incidentItems,
          traingingItems: traingingItems,
          leaveItems: leaveItems,
        }

console.log(data);

        await this.generateToExcel(data)
      },
      async generateToExcel(data) {
        // const Excel = require('exceljs');

        // Create workbook & add worksheet
        const workbook = new Excel.Workbook();
        const worksheet = workbook.addWorksheet('Details');

        const allBorder = {
          top: {
            style: 'thin'
          },
          left: {
            style: 'thin'
          },
          bottom: {
            style: 'thin'
          },
          right: {
            style: 'thin'
          }
        }

        //  color //
        //#region color
        worksheet.getCell('I1').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'Bce5ae'
          },
        };
        worksheet.getCell('A3').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'bfbfbf'
          },
        };
        worksheet.getCell('A6').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'fbf09b'
          },
        };
        worksheet.getCell('I6').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'ffcc99'
          },
        };
        worksheet.getCell('J6').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '9e355b'
          },
        };
        worksheet.getCell('J8').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '9e355b'
          },
        };
        worksheet.getCell('J9').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '9e355b'
          },
        };
        worksheet.getCell('J10').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '9e355b'
          },
        };
        worksheet.getCell('A7').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'd9e2f3'
          },
        };
        worksheet.getCell('I7').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'f2f2f2'
          },
        };
        worksheet.getCell('A8').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'd9e2f3'
          },
        };
        worksheet.getCell('I8').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'f2f2f2'
          },
        };
        worksheet.addConditionalFormatting({
          ref: 'A9:H10',
          rules: [{
            type: 'expression',
            formulae: ['MOD(ROW()+COLUMN(),1)=0'],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                bgColor: {
                  argb: 'ff99cc'
                }
              }
            },
          }]
        })
        worksheet.getCell('I9').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'fbe4d5'
          },
        };
        worksheet.getCell('I10').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'fbe4d5'
          },
        };
        worksheet.getCell('J9').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '993366'
          },
        };
        worksheet.addConditionalFormatting({
          ref: 'A23:I24',
          rules: [{
            type: 'expression',
            formulae: ['MOD(ROW()+COLUMN(),1)=0'],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                bgColor: {
                  argb: 'ff99cc'
                }
              }
            },
          }]
        })
        worksheet.getCell('J23').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '993366'
          },
        };
        worksheet.addConditionalFormatting({
          ref: 'A38:I39',
          rules: [{
            type: 'expression',
            formulae: ['MOD(ROW()+COLUMN(),1)=0'],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                bgColor: {
                  argb: 'ff99cc'
                }
              }
            },
          }]
        })
        worksheet.getCell('J38').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '993366'
          },
        };
        worksheet.addConditionalFormatting({
          ref: 'A46:I47',
          rules: [{
            type: 'expression',
            formulae: ['MOD(ROW()+COLUMN(),1)=0'],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                bgColor: {
                  argb: 'ff99cc'
                }
              }
            },
          }]
        })
        worksheet.getCell('J46').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '993366'
          },
        };
        worksheet.addConditionalFormatting({
          ref: 'A53:I54',
          rules: [{
            type: 'expression',
            formulae: ['MOD(ROW()+COLUMN(),1)=0'],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                bgColor: {
                  argb: 'ff99cc'
                }
              }
            },
          }]
        })
        worksheet.getCell('J53').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '993366'
          },
        };
        worksheet.getCell('A59').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'ee0b11'
          },
        };
        worksheet.getCell('I59').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'ff99cc'
          },
        };
        worksheet.getCell('J59').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'ee0b11'
          },
        };

        //#endregion
        // ===== format font ====
        //#region format font

        worksheet.getCell('A1').alignment = {
          vertical: 'middle',
          horizontal: 'left'
        };
        worksheet.getCell('I1').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('A10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('B10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('C10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('D10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('E10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('F10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('G10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('H10').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('A24').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('B24').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('D24').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('A39').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('B39').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('C39').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('D39').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('A47').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('B47').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('D47').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('A54').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('B54').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        // worksheet.getCell('A').alignment = { vertical: 'middle', horizontal: 'center' };
        // worksheet.getCell('B').alignment = { vertical: 'middle', horizontal: 'center' };
        // worksheet.getCell('C').alignment = { vertical: 'middle', horizontal: 'center' };
        // worksheet.getCell('D').alignment = { vertical: 'middle', horizontal: 'center' };
        // worksheet.getCell('E').alignment = { vertical: 'middle', horizontal: 'center' };
        // worksheet.getCell('F').alignment = { vertical: 'middle', horizontal: 'center' };
        // worksheet.getCell('G').alignment = { vertical: 'middle', horizontal: 'center' };
        // worksheet.getCell('H').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('J6').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        worksheet.getCell('J6').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('J9').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        worksheet.getCell('J9').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('J23').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        worksheet.getCell('J23').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('J38').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        worksheet.getCell('J38').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('J46').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        worksheet.getCell('J46').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('J53').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        worksheet.getCell('J53').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('J59').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        worksheet.getCell('J59').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('I9').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('I23').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('I38').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('I46').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('I53').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };
        worksheet.getCell('I59').alignment = {
          vertical: 'middle',
          horizontal: 'center'
        };

        worksheet.getCell('A59').font = {
          color: {
            argb: 'FFFFFF'
          },
        };
        //#endregion
        //////////////////////  Border ///////////////////////////////////////
        //#region border
        worksheet.getCell('A4').border = allBorder;
        worksheet.getCell('A5').border = allBorder;
        worksheet.getCell('F5').border = allBorder;
        worksheet.getCell('A6').border = allBorder;
        worksheet.getCell('A7').border = allBorder;
        worksheet.getCell('A8').border = allBorder;
        worksheet.getCell('I7').border = allBorder;
        worksheet.getCell('I8').border = allBorder;

        worksheet.getCell('B24').border = allBorder;
        worksheet.getCell('B39').border = allBorder;
        worksheet.getCell('B39').border = allBorder;
        worksheet.getCell('B47').border = allBorder;
        worksheet.getCell('B54').border = allBorder;

        // worksheet.columns.forEach((column, index) => {
        //   column.border = allBorder;
        // });
        worksheet.eachRow({
          includeEmpty: true
        }, function (row, rowNumber) {
          row.eachCell({
            includeEmpty: true
          }, function (cell, colNumber) {
            cell.border = allBorder
          });
        });


        let start = 11
        for (let i = 0; i < 13; i++) {
          worksheet.getCell(`A${start+i}`).border = allBorder
          worksheet.getCell(`B${start+i}`).border = allBorder
          worksheet.getCell(`C${start+i}`).border = allBorder
          worksheet.getCell(`D${start+i}`).border = allBorder
          worksheet.getCell(`E${start+i}`).border = allBorder
          worksheet.getCell(`F${start+i}`).border = allBorder
          worksheet.getCell(`G${start+i}`).border = allBorder
          worksheet.getCell(`H${start+i}`).border = allBorder
          worksheet.getCell(`I${start+i}`).border = allBorder
          worksheet.getCell(`J${start+i}`).border = allBorder
        }

        let start2 = 25
        for (let i = 0; i < 13; i++) {
          worksheet.mergeCells(`B${start2+i}:C${start2+i}`);
          worksheet.mergeCells(`D${start2+i}:H${start2+i}`);

          worksheet.getCell(`A${start2+i}`).border = allBorder
          worksheet.getCell(`B${start2+i}`).border = allBorder
          worksheet.getCell(`C${start2+i}`).border = allBorder
          worksheet.getCell(`D${start2+i}`).border = allBorder
          worksheet.getCell(`E${start2+i}`).border = allBorder
          worksheet.getCell(`F${start2+i}`).border = allBorder
          worksheet.getCell(`G${start2+i}`).border = allBorder
          worksheet.getCell(`H${start2+i}`).border = allBorder
          worksheet.getCell(`I${start2+i}`).border = allBorder
          worksheet.getCell(`J${start2+i}`).border = allBorder
        }

        let start3 = 40
        for (let i = 0; i < 6; i++) {
          worksheet.mergeCells(`D${start3+i}:H${start3+i}`);

          worksheet.getCell(`A${start3+i}`).border = allBorder
          worksheet.getCell(`B${start3+i}`).border = allBorder
          worksheet.getCell(`C${start3+i}`).border = allBorder
          worksheet.getCell(`D${start3+i}`).border = allBorder
          worksheet.getCell(`E${start3+i}`).border = allBorder
          worksheet.getCell(`F${start3+i}`).border = allBorder
          worksheet.getCell(`G${start3+i}`).border = allBorder
          worksheet.getCell(`H${start3+i}`).border = allBorder
          worksheet.getCell(`I${start3+i}`).border = allBorder
          worksheet.getCell(`J${start3+i}`).border = allBorder
        }

        let start4 = 48
        for (let i = 0; i < 5; i++) {
          worksheet.mergeCells(`B${start4+i}:C${start4+i}`);
          worksheet.mergeCells(`D${start4+i}:H${start4+i}`);
          worksheet.getCell(`A${start4+i}`).border = allBorder
          worksheet.getCell(`B${start4+i}`).border = allBorder
          worksheet.getCell(`C${start4+i}`).border = allBorder
          worksheet.getCell(`D${start4+i}`).border = allBorder
          worksheet.getCell(`E${start4+i}`).border = allBorder
          worksheet.getCell(`F${start4+i}`).border = allBorder
          worksheet.getCell(`G${start4+i}`).border = allBorder
          worksheet.getCell(`H${start4+i}`).border = allBorder
          worksheet.getCell(`I${start4+i}`).border = allBorder
          worksheet.getCell(`J${start4+i}`).border = allBorder
        }
        let start5 = 55
        for (let i = 0; i < 4; i++) {
          worksheet.mergeCells(`B${start5+i}:H${start5+i}`);
          worksheet.getCell(`A${start5+i}`).border = allBorder
          worksheet.getCell(`B${start5+i}`).border = allBorder
          worksheet.getCell(`C${start5+i}`).border = allBorder
          worksheet.getCell(`D${start5+i}`).border = allBorder
          worksheet.getCell(`E${start5+i}`).border = allBorder
          worksheet.getCell(`F${start5+i}`).border = allBorder
          worksheet.getCell(`G${start5+i}`).border = allBorder
          worksheet.getCell(`H${start5+i}`).border = allBorder
          worksheet.getCell(`I${start5+i}`).border = allBorder
          worksheet.getCell(`J${start5+i}`).border = allBorder
        }

        //#endregion border
        ///////////////////////////////////////////////////////////

        let textCenter = {
              vertical: 'middle',
              horizontal: 'center'
            }

        worksheet.mergeCells('A1:B2');
        worksheet.mergeCells('C1:H1');
        worksheet.mergeCells('I1:J2');
        worksheet.getCell('A1').value = 'Company:';
        worksheet.getCell('C1').value = 'RiverPark Consultant Company Limited';
        worksheet.getCell('I1').value = data.date;

        worksheet.addRow()
        worksheet.mergeCells('C2:H2');

        worksheet.mergeCells('A3:J3');

        worksheet.mergeCells('A4:J4');
        worksheet.getCell('A4').value = 'Time Sheet Daily Report';

        worksheet.mergeCells('A5:E5');
        worksheet.mergeCells('F5:J5');
        worksheet.getCell('A5').value = "Name : " + data.fullName
        worksheet.getCell('F5').value = "Position : " + data.position

        worksheet.mergeCells('A6:H6');
        worksheet.mergeCells('J6:J8');
        worksheet.getCell('A6').value = "Description"
        worksheet.getCell('I6').value = ""
        worksheet.getCell('J6').value = "Summary"

        worksheet.mergeCells('A7:H7');
        worksheet.getCell('A7').value = "Time In"
        worksheet.getCell('I7').value = data.timeIn

        worksheet.mergeCells('A8:H8');
        worksheet.getCell('A8').value = "Time Out"
        worksheet.getCell('I8').value = data.timeOut


        worksheet.mergeCells('A9:H9');
        worksheet.mergeCells('J9:J10');
        worksheet.getCell('A9').value = "Total Hours for Project"
        worksheet.getCell('I9').value = data.totalHoursforProject
        worksheet.getCell('J9').value = data.totalHoursforProjectText

        worksheet.getCell('A10').value = "No"
        worksheet.getCell('B10').value = "RFC"
        worksheet.getCell('C10').value = "Project Name"
        worksheet.getCell('D10').value = "Project Stage"
        worksheet.getCell('E10').value = "Delivered to"
        worksheet.getCell('F10').value = "STATUS"
        worksheet.getCell('G10').value = "NOTES"
        worksheet.getCell('H10').value = "Task description"

        // loop 13 rows or more than
        for (let index = 0; index < 13; index++) {
          if(index < data.projectItems.length){
            let item = data.projectItems[index]
            worksheet.getCell(`A${11 + index}`).value = item.data[0]
            worksheet.getCell(`B${11 + index}`).value = item.data[1]
            worksheet.getCell(`C${11 + index}`).value = item.data[2]
            worksheet.getCell(`D${11 + index}`).value = item.data[3]
            worksheet.getCell(`E${11 + index}`).value = item.data[4]
            worksheet.getCell(`F${11 + index}`).value = item.data[5]
            worksheet.getCell(`G${11 + index}`).value = item.data[6]
            worksheet.getCell(`H${11 + index}`).value = item.data[7]
            worksheet.getCell(`I${11 + index}`).value = item.data[8]
            worksheet.getCell(`J${11 + index}`).value = item.data[9]

            // worksheet.getCell(`A${11 + index}`).alignment = textCenter
            // worksheet.getCell(`B${11 + index}`).alignment = textCenter
            // worksheet.getCell(`C${11 + index}`).alignment = textCenter
            // worksheet.getCell(`D${11 + index}`).alignment = textCenter
            // worksheet.getCell(`E${11 + index}`).alignment = textCenter
            // worksheet.getCell(`F${11 + index}`).alignment = textCenter
            // worksheet.getCell(`G${11 + index}`).alignment = textCenter
            // worksheet.getCell(`H${11 + index}`).alignment = textCenter
            worksheet.getCell(`I${11 + index}`).alignment = textCenter
            worksheet.getCell(`J${11 + index}`).alignment = textCenter
          }
        }

        worksheet.mergeCells('A23:H23');
        worksheet.mergeCells('J23:J24');
        worksheet.getCell('A23').value = "Total Hours for Non Project"
        worksheet.getCell('I23').value = data.totalHoursforNonProject
        worksheet.getCell('J23').value = data.totalHoursforNonProjectText

        //------------------------------------------------------------

        worksheet.mergeCells('B24:C24');
        worksheet.mergeCells('D24:H24');
        worksheet.getCell('A24').value = "No"
        worksheet.getCell('B24').value = "Non Project"
        worksheet.getCell('D24').value = "Task description"

        //loop

        worksheet.mergeCells('A38:H38');
        worksheet.mergeCells('J38:J39');
        worksheet.mergeCells('D39:I39');
        worksheet.getCell('A38').value = "Total Hours for Incident/Bug Fix"
        worksheet.getCell('I38').value = data.totalHoursforIncident
        worksheet.getCell('J38').value = data.totalHoursforIncidentText

        //------------------------------------------------------------

        worksheet.getCell('A39').value = "No"
        worksheet.getCell('B39').value = "Incident no."
        worksheet.getCell('C39').value = "Incident/Bug fix"
        worksheet.getCell('D39').value = "Task description"

        // loop 


        worksheet.mergeCells('A46:H46');
        worksheet.mergeCells('J46:J47');
        worksheet.mergeCells('D47:I47');
        worksheet.getCell('A46').value = "Total Hours for Training/Seminar"
        worksheet.getCell('I46').value = data.totalHoursforTraining
        worksheet.getCell('J46').value = data.totalHoursforTrainingText

        //------------------------------------------------------------


        worksheet.mergeCells('B47:C47');
        worksheet.getCell('A47').value = "No"
        worksheet.getCell('B47').value = "Training/Seminar"
        worksheet.getCell('D47').value = "Course Name/Description"

        // loop 

        worksheet.mergeCells('A53:H53');
        worksheet.mergeCells('J53:J54');
        worksheet.getCell('A53').value = "Total Hours for Leave"
        worksheet.getCell('I53').value = data.totalHoursforLeave
        worksheet.getCell('J53').value = data.totalHoursforLeaveText

        //------------------------------------------------------------

        worksheet.mergeCells('B54:H54');
        worksheet.getCell('A54').value = "No"
        worksheet.getCell('B54').value = "Leave"

        //  loop 

        worksheet.mergeCells('A59:H59');
        worksheet.getCell('A59').value = "Total Hours"
        worksheet.getCell('I59').value = data.totalHours
        worksheet.getCell('J59').value = data.totalHoursText


        //--------------------------------------------------------------
        // import {
        //   saveAs
        // } from "file-saver";

        const buffer = await workbook.xlsx.writeBuffer();
        const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const fileExtension = '.xlsx';

        const blob = new Blob([buffer], {
          type: fileType
        });

        saveAs(blob, 'Time Sheet 01_03_2566 FristName' + fileExtension);

        // save workbook to disk
        // workbook
        //   .xlsx
        //   .writeFile('sample.xlsx')
        //   .then(() => {
        //     console.log("saved");
        //   })
        //   .catch((err) => {
        //     console.log("err", err);
        //   });

      }
    }
  }

</script>
