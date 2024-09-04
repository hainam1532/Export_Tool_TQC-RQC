import { is } from '@electron-toolkit/utils'
import oracledb from 'oracledb'
const path = require('path')

const pathInstanceClient = is.dev
  ? path.join(__dirname, '/../../resources/instantclient_21_11')
  : path.join(__dirname, '/../../../app.asar.unpacked/resources/instantclient_21_11')

oracledb.outFormat = oracledb.OUT_FORMAT_OBJECT
oracledb.initOracleClient({ libDir: pathInstanceClient })

export async function query(sql) {
  const connection = await oracledb.getConnection({
    user: 'mes00',
    password: 'dbmes00',
    connectionString: '10.30.3.51:1521/APHMES'
  })

  const result = await connection.execute(sql)

  console.log(result)
  return result.rows
}
