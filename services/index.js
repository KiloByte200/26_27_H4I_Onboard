"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getDisasters = getDisasters;
var tslib_1 = require("tslib");
var sp_http_1 = require("@microsoft/sp-http");
function getDisasters(context) {
    return tslib_1.__awaiter(this, void 0, void 0, function () {
        var url, response, data;
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (Math.random() < 0.5) {
                        throw new Error("Simulated Api Failure");
                    }
                    url = "".concat(context.pageContext.web.absoluteUrl, "/_api/web/lists");
                    return [4 /*yield*/, context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1)];
                case 1:
                    response = _a.sent();
                    return [4 /*yield*/, response.json()];
                case 2:
                    data = _a.sent();
                    return [2 /*return*/, data];
            }
        });
    });
}
//# sourceMappingURL=index.js.map