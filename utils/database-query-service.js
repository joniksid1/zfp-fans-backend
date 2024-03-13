const { fanDataDb, priceDb } = require('./db');
const { NotFoundError } = require('./errors/not-found-error');

const { MYSQL_FAN_DATABASE } = process.env;
const { fanModels } = require('../constants/fan-models');

// Получение данных моделей вентиляторов из БД (для fan-data.js)
async function getFanModels() {
  let connection;
  try {
    connection = await fanDataDb.getConnection();
    const [allModelsQuery] = await connection.execute(`
      SELECT DISTINCT model
      FROM ${MYSQL_FAN_DATABASE}.zfr_data;
    `);
    if (allModelsQuery.length === 0) {
      throw new NotFoundError({ message: 'Не удалось найти данные названий вентиляторов в базе' });
    }
    const modelsArray = allModelsQuery.map((result) => result.model);
    return modelsArray;
  } finally {
    if (connection) {
      await connection.release();
    }
  }
}

// Получение данных о точках графиков вентиляторов (для fan-data.js)
async function getFanDataPoints() {
  let connection;
  const fanDataResults = [];

  try {
    const results = await Promise.allSettled(fanModels.map(async (fanModel) => {
      try {
        connection = await fanDataDb.getConnection();
        const [fanDataQuery] = await connection.execute(`
          SELECT x, y
          FROM ${MYSQL_FAN_DATABASE}.${fanModel}_dataset;
        `);
        if (fanDataQuery.length === 0) {
          throw new NotFoundError({ message: `Не удалось найти данные вентилятора ${fanModel} в базе` });
        }
        fanDataResults.push({
          model: fanModel,
          data: fanDataQuery.map((result) => ({ x: result.x, y: result.y })),
        });
      } finally {
        if (connection) {
          await connection.release();
        }
      }
    }));
    const errors = results.filter((result) => result.status === 'rejected').map((result) => result.reason);
    if (errors.length > 0) {
      throw errors[0];
    }
    return fanDataResults;
  } finally {
    if (connection) {
      await connection.release();
    }
  }
}

// Получение данных по ценам и названиям из БД для коммерческого предложения
async function fetchDataQueries(selectedData) {
  const queryResults = await Promise.all(selectedData.map(async (data) => {
    const optionsQuery = await fanDataDb.query(`
      SELECT ZRS, ZRSI, ZRN, ZRF, ZRC, ZRD, Regulator
      FROM ${MYSQL_FAN_DATABASE}.zfr_options
      WHERE model = ?
    `, [data.fanName]);

    const [priceDbData] = await priceDb.query(`
      SELECT *
      FROM Price
      WHERE Model IN (?, ?, ?, ?, ?, ?, ?, ?);
    `, [
      data.fanName,
      optionsQuery[0][0].ZRS,
      optionsQuery[0][0].ZRSI,
      optionsQuery[0][0].ZRN,
      optionsQuery[0][0].ZRF,
      optionsQuery[0][0].ZRC,
      optionsQuery[0][0].ZRD,
      optionsQuery[0][0].Regulator,
    ]);

    if (priceDbData.length === 0 || optionsQuery.length === 0) {
      throw new NotFoundError({ message: 'Не удалось найти данные в базе' });
    }

    return { data, priceDbData, optionsQuery };
  }));

  return queryResults;
}

// Получение технических характеристик вентиляторов
async function getFanTechnicalData(fanName) {
  const [techDataQuery] = await fanDataDb.query(`
    SELECT
      id,
      model,
      max_airflow_m3h,
      max_static_pressure_pa,
      voltage_V,
      power_consumption_kW,
      max_operating_current_A,
      rotation_frequency_rpm,
      sound_power_level_dBA,
      airflow_temperature_range,
      capacitor_mF,
      electrical_connections_scheme
    FROM ${MYSQL_FAN_DATABASE}.zfr_data
    WHERE model = ?
  `, [fanName]);

  if (techDataQuery.length === 0) {
    throw new NotFoundError({ message: 'Не удалось найти данные технических характеристик вентиляторов в базе' });
  }

  return techDataQuery[0];
}

// Получение габаритов вентиляторов
async function getFanDimensionsData(fanName) {
  const [dimensionsQuery] = await fanDataDb.query(`
    SELECT
      id,
      model,
      l,
      l1,
      l2,
      h,
      d,
      l3,
      kg
    FROM ${MYSQL_FAN_DATABASE}.zfr_dimensions
    WHERE model = ?
  `, [fanName]);

  if (dimensionsQuery.length === 0) {
    throw new NotFoundError({ message: 'Не удалось найти данные размеров вентиляторов в базе' });
  }

  return dimensionsQuery[0];
}

// Получение данных названий опций
async function getFanOptionsName(fanName) {
  const [optionsQuery] = await fanDataDb.query(`
    SELECT ZRS, ZRSI, ZRN, ZRF, ZRC, ZRD
    FROM ${MYSQL_FAN_DATABASE}.zfr_options
    WHERE model = ?
  `, [fanName]);

  if (optionsQuery.length === 0) {
    throw new NotFoundError({ message: 'Не удалось найти названия опций в базе' });
  }

  return optionsQuery[0];
}

// Получение габаритов монтажных стаканов
async function getSocketDimensionsData(options) {
  const [socketDimensionsQuery] = await fanDataDb.query(`
    SELECT
      id,
      TypeSize,
      Model,
      Hole_Spacing_D,
      outer_socket_width_E,
      Thread_Type_M,
      inner_socket_width_G,
      outer_platform_width_F,
      height_H,
      Weight_kg
    FROM ${MYSQL_FAN_DATABASE}.zrs_zrsi_zrn_dimensions
    WHERE Model IN (?, ?, ?)
  `, [options.ZRS, options.ZRSI, options.ZRN]);

  if (socketDimensionsQuery.length === 0) {
    throw new NotFoundError({ message: 'Данные по монтажным стаканам не найдены в базе' });
  }

  return socketDimensionsQuery;
}

// Получение габаритов опций "гибкая вставка", "фланец", "обратный клапан"
async function getOtherOptionsDimensionsData(options) {
  const [otherOptionsDimensionsQuery] = await fanDataDb.query(`
    SELECT
      id,
      TypeSize,
      Model,
      Inner_Diameter_d,
      Middle_Diameter_e,
      Inner_Diameter_corrected_D,
      Height_h,
      Length_L,
      Diameter_D2,
      Weight_kg
    FROM ${MYSQL_FAN_DATABASE}.zrd_zrc_zrf_dimensions
    WHERE Model IN (?, ?, ?)
  `, [options.ZRD, options.ZRC, options.ZRF]);

  if (otherOptionsDimensionsQuery.length === 0) {
    throw new NotFoundError({ message: 'Данные по опциям "фланец", "гибкая вставка", "обратный клапан" не найдены в базе' });
  }

  return otherOptionsDimensionsQuery;
}

module.exports = {
  getFanModels,
  getFanDataPoints,
  fetchDataQueries,
  getFanTechnicalData,
  getFanDimensionsData,
  getFanOptionsName,
  getSocketDimensionsData,
  getOtherOptionsDimensionsData,
};
