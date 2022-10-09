const axios = require(`axios`).default;

exports.getData = async function (method, url, headers) {
  try {
    const response = await axios({
      method,
      url,
      headers,
    });
    const respData = await response.data.response;
    return respData;
  } catch (error) {
    console.log(error);
  }
};

exports.sendData = async function (method, url, headers, data) {
  try {
    const response = await axios({
      method,
      url,
      headers,
      data,
    });
    const respData = await response.data.response;
    return respData;
  } catch (error) {
    console.log(error);
  }
};

exports.downloadData = async function (method, url, headers, responseType) {
  try {
    const response = await axios({
      method,
      url,
      headers,
      responseType,
    });
    const respData = await response.data;
    return respData;
  } catch (error) {
    console.log(error);
  }
};
