import { useState, useEffect } from 'react'
import { Button, Flex, Card, DatePicker, Select } from 'antd'
import { toast } from 'react-hot-toast'
import exportExcel from '../../main/module/exportExcel'
import picture1 from '../../../resources/Browser stats-bro.svg'
import picture2 from '../../../resources/Data report-amico.svg'

const { RangePicker } = DatePicker

function App() {
  const [dataTQC, setDataTQC] = useState([])
  const [dataRQC, setDataRQC] = useState([])
  //starttime
  const [startTime, setStartTime] = useState(null)
  const [startTimeRQC, setStartTimeRQC] = useState(null)
  //endtime
  const [endTime, setEndTime] = useState(null)
  const [endTimeRQC, setEndTimeRQC] = useState(null)
  //select status
  const [selectedStatus, setSelectedStatus] = useState('')
  const [selectedStatusRQC, setSelectedStatusRQC] = useState('')

  const onChangeTime = (dates, dateStrings) => {
    if (dates) {
      setStartTime(dateStrings[0])
      setEndTime(dateStrings[1])
    } else {
      setStartTime(null)
      setEndTime(null)
    }
    //console.log(dates);
  }

  const onChangeTimeRQC = (dates, dateStrings) => {
    if (dates) {
      setStartTimeRQC(dateStrings[0])
      setEndTimeRQC(dateStrings[1])
    } else {
      setStartTimeRQC(null)
      setEndTimeRQC(null)
    }
    //console.log(startTime,endTime);
  }

  const onChange = (value) => {
    setSelectedStatus(value)
    //console.log(value);
  }

  const onChangeRQC = (value) => {
    setSelectedStatusRQC(value)
  }

  const onSearch = (value) => {
    console.log('search:', value)
  }

  const onSearchRQC = (value) => {
    console.log('search:', value)
  }

  //TQC
  const handleExportExcelTQC = async () => {
    try {
      //const result = await executeDatabaseQuery();

      const exportData = dataTQC.map((item) => ({
        task_no: item.TASK_NO,
        task_state: item.TASK_STATE,
        createdate: item.CREATEDATE,
        mer_po: item.MER_PO,
        se_id: item.SE_ID,
        prod_no: item.PROD_NO,
        status_ship: item.STATUS_SHIP,
        date_ship: item.DATE_SHIP,
        department: item.DEPARTMENT,
        production_line_name: item.PRODUCTION_LINE_NAME,
        shoe_no: item.SHOE_NO,
        name_t: item.NAME_T,
        workshop_section_name: item.WORKSHOP_SECTION_NAME,
        total: item.TOTAL,
        first_qualified_num: item.FIRST_QUALIFIED_NUM,
        first_unqualified_num: item.FIRSTUNQUALIFIEDNUM,
        rftpass: item.RFTPASS,
        top3issue: item.TOP3ISSUE
      }))

      console.log('Data excel TQC: ', exportData)

      exportExcel(`[TQC]-data(${Date.now()}).xlsx`, [
        {
          name: 'Sheet1',
          columns: [
            { header: 'Task No.', key: 'task_no', width: 20 },
            { header: 'Task Status', key: 'task_state', width: 15 },
            { header: 'Created Date', key: 'createdate', width: 20 },
            { header: 'PO', key: 'mer_po', width: 20 },
            { header: 'SO', key: 'se_id', width: 20 },
            { header: 'ART', key: 'prod_no', width: 8 },
            { header: 'Status shipment', key: 'status_ship', width: 20 },
            { header: 'Date shipment', key: 'date_ship', width: 20 },
            { header: 'Department', key: 'department', width: 14 },
            { header: 'Line Name', key: 'production_line_name', width: 15 },
            { header: 'Shoe No.', key: 'shoe_no', width: 19 },
            { header: 'Shoe Name', key: 'name_t', width: 18 },
            { header: 'Section Name', key: 'workshop_section_name', width: 15 },
            { header: 'Total Inspection', key: 'total', width: 15 },
            { header: 'Total Pass', key: 'first_qualified_num', width: 15 },
            { header: 'Total Fail', key: 'first_unqualified_num', width: 15 },
            { header: 'RFT %', key: 'rftpass', width: 15 },
            { header: 'Top Issue', key: 'top3issue', width: 50 }
          ],
          data: exportData,
          callback: ({ worksheet }) => {
            const col = 18
            const row = 1000

            // Style for header
            worksheet.getRow(1).height = 42.75
            Array.from({ length: col }).forEach((_, index) => {
              const cell = worksheet.getRow(1).getCell(1 + index)

              // Border
              cell.border = {
                top: { style: 'thick', color: { argb: '000000' } },
                left: {
                  style: index === 0 ? 'thick' : 'thin',
                  color: { argb: '000000' }
                },
                bottom: { style: 'thin', color: { argb: '000000' } },
                right: {
                  style: index === col - 1 ? 'thick' : 'thin',
                  color: { argb: '000000' }
                }
              }

              // Alignment
              cell.alignment = {
                vertical: 'middle',
                horizontal: 'center',
                wrapText: true
              }

              // Background color
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD080' }
              }
            })

            // Style for rows
            Array.from({ length: row }).forEach((_, indexRow) => {
              Array.from({ length: col }).forEach((_, indexColumn) => {
                const cell = worksheet.getRow(2 + indexRow).getCell(1 + indexColumn)

                // Border
                cell.border = {
                  top: { style: 'thin', color: { argb: '000000' } },
                  left: {
                    style: indexColumn === 0 ? 'thick' : 'thin',
                    color: { argb: '000000' }
                  },
                  bottom: {
                    style: indexRow === row - 1 ? 'thick' : 'thin',
                    color: { argb: '000000' }
                  },
                  right: {
                    style: indexColumn === col - 1 ? 'thick' : 'thin',
                    color: { argb: '000000' }
                  }
                }

                // Alignment and text wrapping
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center',
                  wrapText: true
                }

                // Background color for alternate rows
                if (indexRow % 2 !== 0) {
                  cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'F3F3F3' }
                  }
                }

                // Set cell format to text
                cell.numFmt = '@'
              })
            })
          },
          views: [{ showGridLines: false }]
        }
      ])

      toast.success('Export successful!', {
        duration: 5000
      })
    } catch (error) {
      toast.error('Export failed. Please try again.')
      console.error('Export error:', error)
    }
  }

  const executeDatabaseQuery = async () => {
    if (!startTime || !endTime || !selectedStatus) {
      toast('Vui lòng chọn ngày và trạng thái', {
        icon: '⚠️',
        style: {
          border: '1px solid #713200',
          padding: '6px',
          color: '#713200'
        },
        duration: 2000
      })
      console.error('Date and status are required.')
      return
    }

    try {
      // Lọc theo date
      let timeFilter2 = ''
      if (startTime && endTime) {
        timeFilter2 = `t.createdate BETWEEN '${startTime}' AND '${endTime}'`
      }

      // Lọc theo status
      let statusFilter = ''
      if (selectedStatus) {
        statusFilter = `AND t.task_state IN (${selectedStatus})`
      }

      const sqlQuery = `WITH tj AS
 (SELECT t.prod_no,
         t.department,
         t.production_line_code,
         t.workshop_section_no,
         SUM(CASE
               WHEN b.commit_type = '0' THEN
                1
               ELSE
                0
             END) AS hg,
         SUM(CASE
               WHEN b.commit_type = '1' THEN
                1
               ELSE
                0
             END) AS bhg,
         SUM(CASE
               WHEN b.commit_type = '2' THEN
                1
               ELSE
                0
             END) AS fxhg,
         SUM(CASE
               WHEN b.commit_type = '3' THEN
                1
               ELSE
                0
             END) AS fxbhg,
         SUM(CASE
               WHEN b.commit_type = '4' THEN
                1
               ELSE
                0
             END) AS bp
    FROM tqc_task_m t
    LEFT JOIN tqc_task_commit_m b
      ON t.task_no = b.task_no
    LEFT JOIN base005m c
      ON t.production_line_code = c.department_code
   WHERE ${timeFilter2} ${statusFilter}
   --WHERE t.task_no = 'T2024090373'
   GROUP BY t.department,
            t.production_line_code,
            t.workshop_section_no,
            t.prod_no),
tm AS
 (SELECT t.task_no,
         t.prod_no,
         t.shoe_no,
         s.name_t,
         t.createdate AS createdate,
         t.workshop_section_no,
         t.department,
         c.department_name AS department_name,
         CASE
           WHEN t.task_state = '0' THEN
            'Đang tiến hành'
           WHEN t.task_state = '1' THEN
            'Đã hoàn thành'
           WHEN t.task_state = '2' THEN
            'Kết thúc'
           WHEN t.task_state = '3' THEN
            'Đang mở hộp'
           WHEN t.task_state = '4' THEN
            'Đang kiểm tra và tạm dừng'
           WHEN t.task_state = '5' THEN
            'Kiểm tra kết thúc'
         END AS task_state,
         c.udf05,
         t.production_line_code,
         t.mer_po,
         t.se_id,
         shh.posting_date,
         CASE
           WHEN shh.status = '7' THEN
            'Đã xuất hàng'
           WHEN shh.status IS NULL THEN
            'Chưa xuất hàng'
         END AS status_ship
    FROM tqc_task_m t
    LEFT JOIN base005m c
      ON t.production_line_code = c.department_code
    LEFT JOIN bdm_rd_style s
      ON t.shoe_no = s.shoe_no
    LEFT JOIN bmd_se_shipment_m shh
      ON t.mer_po = shh.po_no
  WHERE ${timeFilter2} ${statusFilter}
  --WHERE t.task_no = 'T2024090373'
  ),
task_detail AS
 (SELECT inspection_name, commit_type
    FROM tqc_task_detail_t t
    LEFT JOIN tqc_task_item_c c
      ON t.union_id = c.id),
tuojiaoNum AS
 (SELECT COUNT(1) AS tuojiao_num
    FROM task_detail
   WHERE inspection_name = '脱胶'
     AND commit_type != 0),
InspectionCounts AS
 (SELECT a.task_no,
           b.inspection_name,
           COUNT(b.inspection_name) AS count_name,
           COUNT(CASE WHEN a.commit_type IS NOT NULL THEN 1 ELSE NULL END) AS commit_type_count
    FROM tqc_task_detail_t a
    LEFT JOIN tqc_task_item_c b ON a.union_id = b.id
    LEFT JOIN tqc_task_m dm ON a.task_no = dm.task_no
    GROUP BY a.task_no, b.inspection_name),
RankedIssues AS
 (SELECT task_no,
           inspection_name,
           count_name,
           commit_type_count,
           ROW_NUMBER() OVER (PARTITION BY task_no ORDER BY count_name DESC, inspection_name) AS rn
    FROM InspectionCounts),
top3issue AS
 (SELECT task_no,
       LISTAGG(inspection_name || ' (' || commit_type_count || ')', ', ') WITHIN GROUP (ORDER BY count_name DESC, inspection_name) AS top3issue
FROM RankedIssues
WHERE rn <= 3
GROUP BY task_no)
SELECT tm.udf05,
       LISTAGG(DISTINCT tm.task_no, ',') WITHIN GROUP(ORDER BY tm.task_no) AS task_no,
       tm.prod_no,
       tm.department,
       tm.department_name AS production_line_name,
       tm.production_line_code AS line_name,
       '' AS problems,
       COALESCE((SELECT tuojiao_num FROM tuojiaoNum), 0) AS tuojiao_nums,
       '' AS tuojiao_rate,
       MAX(tm.shoe_no) AS shoe_no,
       MAX(tm.name_t) AS name_t,
       (SELECT gd.workshop_section_name
          FROM bdm_workshop_section_m gd
         WHERE gd.workshop_section_no = tm.workshop_section_no) AS workshop_section_name,
       MAX(COALESCE(tj.hg + tj.bhg + tj.fxhg + tj.fxbhg + tj.bp, 0)) AS total,
       MAX(COALESCE(tj.hg + tj.bhg, 0)) AS firstCheckTotal,
       MAX(COALESCE(tj.hg, 0)) AS first_qualified_num,
       MAX(COALESCE(tj.bhg, 0)) AS firstUnqualifiedNum,
       MAX(COALESCE(tj.hg + tj.fxhg, 0)) AS qualified,
       MAX(COALESCE(tj.bp, 0)) AS bnum,
       MAX(COALESCE(tj.fxhg, 0)) AS returnFixPass,
       MAX(COALESCE(tj.fxbhg, 0)) AS returnFixNoPass,
       MAX(COALESCE(tj.fxhg + tj.fxbhg, 0)) AS returnFixSum,
       (CASE
         WHEN COALESCE(SUM(tj.hg + tj.fxhg), 0) != 0 THEN
          ROUND(COALESCE(SUM(tj.hg + tj.fxhg), 0) /
                COALESCE(SUM(tj.hg + tj.bhg + tj.fxhg + tj.fxbhg + tj.bp), 0) * 100,
                2) || '%'
         ELSE
          '0%'
       END) AS totalpass,
       (CASE
         WHEN COALESCE(SUM(tj.hg + tj.bhg), 0) != 0 THEN
          ROUND(COALESCE(SUM(tj.hg), 0) / COALESCE(SUM(tj.hg + tj.bhg), 0) * 100,
                2) || '%'
         ELSE
          '0%'
       END) AS rftpass,
       (CASE
         WHEN COALESCE(MAX(tj.fxhg + tj.fxbhg), 0) != 0 THEN
          ROUND(COALESCE(MAX(tj.fxhg), 0) /
                COALESCE(MAX(tj.hg + tj.bhg + tj.fxhg + tj.fxbhg + tj.bp), 0) * 100,
                2) || '%'
         ELSE
          '0%'
       END) AS returnRFTRate,
       (MIN(tm.createdate) || '~' || MAX(tm.createdate)) AS datee,
       top3issue.top3issue,
       tm.task_state,
       tm.mer_po,
       tm.se_id,
       TO_CHAR(tm.posting_date, 'yyyy-mm-dd') AS date_ship,
       tm.status_ship,
       MIN(tm.createdate) AS createdate
  FROM tm
  LEFT JOIN tj
    ON tj.prod_no = tm.prod_no
   AND tj.workshop_section_no = tm.workshop_section_no
   AND tj.department = tm.department
   AND tj.production_line_code = tm.production_line_code
  LEFT JOIN top3issue
    ON top3issue.task_no = tm.task_no
 GROUP BY tm.udf05,
          tm.prod_no,
          tm.workshop_section_no,
          tm.department,
          tm.department_name,
          tm.production_line_code,
          tm.task_state,
          tm.mer_po,
          tm.se_id,
          tm.createdate,
          tm.posting_date,
          tm.status_ship,
          top3issue.top3issue
 ORDER BY tm.udf05,
          tm.prod_no,
          tm.workshop_section_no,
          tm.department,
          tm.department_name,
          tm.production_line_code,
          tm.task_state,
          tm.mer_po,
          tm.se_id,
          tm.createdate,
          tm.posting_date,
          tm.status_ship DESC`

      //console.log('SQL Query:', sqlQuery);
      const result = await window.electron.ipcRenderer.invoke('query', sqlQuery)

      setDataTQC(result)
      console.log('Data query TQC: ', result)
      //return result;
    } catch (error) {
      console.error('Error query database:', error)
    }
  }

  //RQC
  const handleExportExcelRQC = async () => {
    try {
      // const result = await executeDatabaseQueryRQC();

      const exportData = dataRQC.map((item) => ({
        task_no: item.TASK_NO,
        task_state: item.TASK_STATE,
        createdate: item.CREATEDATE,
        mer_po: item.MER_PO,
        se_id: item.SO,
        prod_no: item.PROD_NO,
        status_ship: item.STATUS_SHIP,
        date_ship: item.DATE_SHIP,
        department: item.DEPARTMENT,
        production_line_code: item.PRODUCTION_LINE_CODE,
        shoe_no: item.SHOE_NO,
        shoe_name: item.SHOE_NAME,
        workshop_section_name: item.WORKSHOP_SECTION_NAME,
        total_qty: item.TOTAL_QTY,
        pass_qty: item.PASS_QTY,
        bad_qty: item.BAD_QTY,
        qty_percent: item.QTY_PERCENT,
        top3issue: item.TOP3ISSUE
      }))

      console.log('Data excel RQC: ', exportData)

      exportExcel(`[RQC]-data(${Date.now()}).xlsx`, [
        {
          name: 'Sheet1',
          columns: [
            { header: 'Task No.', key: 'task_no', width: 20 },
            { header: 'Task Status', key: 'task_state', width: 15 },
            { header: 'Created Date', key: 'createdate', width: 20 },
            { header: 'PO', key: 'mer_po', width: 20 },
            { header: 'SO', key: 'so', width: 20 },
            { header: 'ART', key: 'prod_no', width: 8 },
            { header: 'Status shipment', key: 'status_ship', width: 20 },
            { header: 'Date shipment', key: 'date_ship', width: 20 },
            { header: 'Department', key: 'department', width: 14 },
            { header: 'Line Code', key: 'production_line_code', width: 15 },
            { header: 'Shoe No.', key: 'shoe_no', width: 19 },
            { header: 'Shoe Name', key: 'shoe_name', width: 18 },
            { header: 'Section Name', key: 'workshop_section_name', width: 15 },
            { header: 'Total Inspection', key: 'total_qty', width: 15 },
            { header: 'Total Pass', key: 'pass_qty', width: 15 },
            { header: 'Total Fail', key: 'bad_qty', width: 15 },
            { header: 'RFT %', key: 'qty_percent', width: 15 },
            { header: 'Top Issue', key: 'top3issue', width: 50 }
          ],
          data: exportData,
          callback: ({ worksheet }) => {
            const col = 18
            const row = 1000

            // Style for header
            worksheet.getRow(1).height = 42.75
            Array.from({ length: col }).forEach((_, index) => {
              const cell = worksheet.getRow(1).getCell(1 + index)

              // Border
              cell.border = {
                top: { style: 'thick', color: { argb: '000000' } },
                left: {
                  style: index === 0 ? 'thick' : 'thin',
                  color: { argb: '000000' }
                },
                bottom: { style: 'thin', color: { argb: '000000' } },
                right: {
                  style: index === col - 1 ? 'thick' : 'thin',
                  color: { argb: '000000' }
                }
              }

              // Alignment
              cell.alignment = {
                vertical: 'middle',
                horizontal: 'center',
                wrapText: true
              }

              // Background color
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD080' }
              }
            })

            // Style for rows
            Array.from({ length: row }).forEach((_, indexRow) => {
              Array.from({ length: col }).forEach((_, indexColumn) => {
                const cell = worksheet.getRow(2 + indexRow).getCell(1 + indexColumn)

                // Border
                cell.border = {
                  top: { style: 'thin', color: { argb: '000000' } },
                  left: {
                    style: indexColumn === 0 ? 'thick' : 'thin',
                    color: { argb: '000000' }
                  },
                  bottom: {
                    style: indexRow === row - 1 ? 'thick' : 'thin',
                    color: { argb: '000000' }
                  },
                  right: {
                    style: indexColumn === col - 1 ? 'thick' : 'thin',
                    color: { argb: '000000' }
                  }
                }

                // Alignment and text wrapping
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center',
                  wrapText: true
                }

                // Background color for alternate rows
                if (indexRow % 2 !== 0) {
                  cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'F3F3F3' }
                  }
                }

                // Set cell format to text
                cell.numFmt = '@'
              })
            })
          },
          views: [{ showGridLines: false }]
        }
      ])

      toast.success('Export successful!', {
        duration: 5000
      })
    } catch (error) {
      toast.error('Export failed. Please try again.')
      console.error('Export error:', error)
    }
  }

  const executeDatabaseQueryRQC = async () => {
    if (!startTimeRQC || !endTimeRQC || !selectedStatusRQC) {
      toast('Vui lòng chọn ngày và trạng thái', {
        icon: '⚠️',
        style: {
          border: '1px solid #713200',
          padding: '6px',
          color: '#713200'
        },
        duration: 2000
      })
      console.error('Date and status are required.')
      return
    }

    try {
      // Lọc theo date
      let timeFilterRQC = ''
      if (startTimeRQC && endTimeRQC) {
        timeFilterRQC = `m.createdate BETWEEN '${startTimeRQC}' AND '${endTimeRQC}'`
      }

      // Lọc theo status
      let statusFilterRQC = ''
      if (selectedStatusRQC) {
        statusFilterRQC = `AND m.task_state IN (${selectedStatusRQC})`
      }

      const sqlQuery = `WITH task_counts AS (
                              SELECT task_no,
                                    COUNT(1) AS total_count,
                                    COUNT(CASE
                                            WHEN commit_type = '0' THEN 1
                                          END) AS commit_zero_count,
                                    SUM(CASE
                                          WHEN commit_type = '0' THEN 1
                                          ELSE 0
                                        END) AS hg,
                                    SUM(CASE
                                          WHEN commit_type = '1' THEN 1
                                          ELSE 0
                                        END) AS bhg,
                                        SUM(CASE
                                          WHEN commit_type IN ('0', '1') THEN 1
                                          ELSE 0
                                        END) AS total_commit_count
                                FROM rqc_task_detail_t
                              GROUP BY task_no
                            ),
                            task_detail AS (
                              SELECT inspection_name, commit_type
                                FROM rqc_task_detail_t t
                                LEFT JOIN rqc_task_item_c c
                                  ON t.task_no = c.task_no
                            ),
                            top3issue AS (
                              SELECT listagg(distinct inspection_name, '/') within group(order by count desc) as top3issue
                                FROM (
                                  SELECT inspection_name, count(1) as count
                                    FROM task_detail
                                  GROUP BY inspection_name
                                  ORDER BY count desc
                                )
                              WHERE rownum <= 3
                            ),
                            order_times AS (
                              SELECT task_no,
                                    CASE
                                      WHEN MODIFYDATE IS NOT NULL THEN
                                        REPLACE(MODIFYDATE, '-', '') || REPLACE(MODIFYTIME, ':', '')
                                      ELSE
                                        REPLACE(CREATEDATE, '-', '') || REPLACE(CREATETIME, ':', '')
                                    END AS order_time
                                FROM rqc_task_m
                            )
                            SELECT a.*,
                                  CASE
                                    WHEN t.total_count > 0 THEN
                                      TO_CHAR(ROUND(t.commit_zero_count / t.total_count, 4) * 100) || '%'
                                    ELSE
                                      '0%'
                                  END AS qty_percent,
                                  t.total_commit_count as total_qty
                              FROM (
                                SELECT m.task_no,
                                      m.workshop_section_no,
                                      m.develop_season,
                                      m.shoe_no,
                                      r.name_t AS shoe_name,
                                      m.prod_no,
                                      m.mer_po,
                                      m.production_line_code,
                                      m.department,
                                      m.se_id,
                                      to_char(shh.posting_date, 'yyyy-mm-dd') as date_ship,
                                      CASE
                                        WHEN shh.status = '7' THEN 'Đã xuất hàng'
                                        WHEN shh.status is null THEN 'Chưa xuất hàng'
                                      END AS status_ship,
                                      (SELECT gd.workshop_section_name
                                          FROM bdm_workshop_section_m gd
                                        WHERE gd.workshop_section_no = m.workshop_section_no) AS workshop_section_name,
                                      (SELECT se_id
                                          FROM bdm_se_order_master
                                        WHERE mer_po = m.mer_po
                                        FETCH FIRST ROW ONLY) AS SO,
                                      m.createdate,
                                      CASE
                                        WHEN task_state = '0' THEN 'Đang tiến hành'
                                        WHEN task_state = '1' THEN 'Dừng lại'
                                        WHEN task_state = '2' THEN 'Kết thúc'
                                      END AS task_state,
                                      CASE
                                        WHEN RESULT = '0' THEN 'PASS'
                                        WHEN RESULT = '1' THEN 'FAIL'
                                      END AS RESULT,
                                      (SELECT COUNT(*)
                                          FROM rqc_task_detail_t d
                                        WHERE d.task_no = m.task_no AND commit_type = '1') AS bad_qty,
                                        (SELECT COUNT(*)
                                          FROM rqc_task_detail_t d
                                        WHERE d.task_no = m.task_no AND commit_type in ('0')) AS pass_qty,
                                      (SELECT top3issue FROM top3issue) as top3issue
                                  FROM rqc_task_m m
                                  LEFT JOIN bdm_rd_style r
                                    ON m.shoe_no = r.shoe_no
                                  LEFT JOIN bmd_se_shipment_m shh
                                    ON m.mer_po = shh.po_no
                                WHERE ${timeFilterRQC} ${statusFilterRQC}
                              ) a
                              LEFT JOIN task_counts t
                                ON a.task_no = t.task_no
                              LEFT JOIN order_times o
                                ON a.task_no = o.task_no
                            ORDER BY o.order_time DESC`
      //console.log('SQL Query:', sqlQuery);
      const result = await window.electron.ipcRenderer.invoke('query', sqlQuery)

      setDataRQC(result)
      console.log('Data query RQC: ', result)

      //return result;
    } catch (error) {
      console.error('Error query database:', error)
    }
  }

  useEffect(() => {
    if (dataRQC.length > 0) {
      handleExportExcelRQC()
    }
  }, [dataRQC])

  useEffect(() => {
    if (dataTQC.length > 0) {
      handleExportExcelTQC()
    }
  }, [dataTQC])

  return (
    <div className="flex flex-1 flex-col rounded p-5 gap-5 overflow-hidden bg-gray-100 h-screen">
      <div className="flex items-center justify-between flex-wrap">
        <div className="font-thin text-[30px]">Export data</div>
      </div>

      <div className="flex gap-8">
        <div className="flex-1 shadow-md">
          <Card title="TQC Data" bordered={false}>
            <div className="flex gap-3">
              <div className="flex-1">
                <RangePicker onChange={onChangeTime} />
              </div>
              <div>
                <Select
                  showSearch
                  placeholder="Select status"
                  optionFilterProp="label"
                  onChange={onChange}
                  allowClear
                  options={[
                    {
                      value: '0',
                      label: 'Open'
                    },
                    {
                      value: '2',
                      label: 'Close'
                    },
                    {
                      value: '0,2',
                      label: 'Both'
                    }
                  ]}
                />
              </div>
            </div>
            <div className="flex-1">
              <img className="h-full w-full" src={picture1} alt="export image" />
            </div>
            <div className="mt-2">
              <Button
                className="bg-blue-400 font-semibold text-white w-full"
                onClick={() => executeDatabaseQuery()}
              >
                Export
              </Button>
            </div>
          </Card>
        </div>

        <div className="flex-1 shadow-md">
          <Card title="RQC Data" bordered={false}>
            <div className="flex gap-3">
              <div className="flex-1">
                <RangePicker onChange={onChangeTimeRQC} />
              </div>
              <div>
                <Select
                  showSearch
                  placeholder="Select status"
                  optionFilterProp="label"
                  onChange={onChangeRQC}
                  onSearch={onSearchRQC}
                  allowClear
                  options={[
                    {
                      value: '0',
                      label: 'Open'
                    },
                    {
                      value: '2',
                      label: 'Close'
                    },
                    {
                      value: '0,2',
                      label: 'Both'
                    }
                  ]}
                />
              </div>
            </div>
            <div className="w-full h-full">
              <img src={picture2} alt="export image" />
            </div>
            <div className="mt-2">
              <Button
                className="bg-blue-400 font-semibold text-white w-full"
                onClick={() => executeDatabaseQueryRQC()}
              >
                Export
              </Button>
            </div>
          </Card>
        </div>
      </div>
    </div>
  )
}

export default App
