import axios from "axios";
import { StaticConst } from "../helper/Const";
export class AsyncHelper {
  ACCESS_TOKEN = "";
  HEADERS = {
    headers: {},
  };
  constructor(token) {
    if (token) {
      this.ACCESS_TOKEN = token;
      this.HEADERS.headers["Authorization"] = this.ACCESS_TOKEN;
    }
  }
  getData = async (url) => {
    if (!url) return;
    return new Promise((resolve, reject) => {
      axios
        .get(StaticConst.graphAPIUrl + url, this.HEADERS)
        .then((res) => {
          resolve(res);
        })
        .catch((ex) => {
          console.error(ex);
          reject(ex);
        });
    });
  };
  postData = async (url, param) => {
    if (!url) return;

    this.HEADERS.headers["Content-Type"] = "application/json";
    //Content-Type , //application/json
    return new Promise((resolve, reject) => {
      axios
        .patch(StaticConst.graphAPIUrl + url, param, this.HEADERS)
        .then((res) => {
          resolve(res);
        })
        .catch((ex) => {
          console.error(ex);
          reject(ex);
        });
    });
  };
  postToAzureFunction = async (param) => {
    this.HEADERS.headers["x-functions-key"] = StaticConst.azureFnMasterKey;
    this.HEADERS.headers["Content-Type"] = "application/json";
    return new Promise((resolve, reject) => {
      axios
        .post(StaticConst.GetAccessTokenFromCodeAzureFnUrl, param, this.HEADERS)
        .then((res) => {
          resolve(res);
        })
        .catch((ex) => {
          console.error(ex);
          reject(ex);
        });
    });
  };
}
