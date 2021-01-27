import * as assert from "assert";
import * as mocha from "mocha";
import { parseNumber } from "office-addin-cli";
import { AppType, runDevServer, startDebugging, stopDebugging } from "office-addin-debugging";
import { toOfficeApp } from "office-addin-manifest";
import * as officeAddinTestHelpers from "office-addin-test-helpers";
import * as officeAddinTestServer from "office-addin-test-server";
import * as path from "path";
import * as testHelpers from "./src/test-helpers";
const hosts = ["Excel", "Word"];
const manifestPath = path.resolve(`${process.cwd()}/test/test-manifest.xml`);
const testServerPort: number = 4201;
const excelDoc: string = "https://microsoft-my.sharepoint-df.com/:x:/r/personal/cowen_microsoft_com/_layouts/15/Doc.aspx?sourcedoc=%7B1CC01E6F-5588-428D-A597-D2E008F3C608%7D&file=Book.xlsx&action=default&mobileredirect=true";
const wordDoc: string = "https://microsoft-my.sharepoint-df.com/:w:/r/personal/cowen_microsoft_com/_layouts/15/Doc.aspx?sourcedoc=%7BE0F19AB3-BDAC-4438-99D3-1729867C3624%7D&file=Document.docx&action=default&mobileredirect=true";
let testServer: officeAddinTestServer.TestServer;
process.env.WEB_SIDELOAD_TEST = "true";

describe(`Test Task Pane Project Add-ins`, function () {
    this.beforeAll(`Start dev-server`, async function () {
        this.timeout(0);
        const devServerCmd = `npm run dev-server -- --config ./test/webpack.config.js`;
        const devServerPort = parseNumber(process.env.npm_package_config_dev_server_port || 3000);
        await runDevServer(devServerCmd, devServerPort);
    }),
    this.afterAll(`Teardown test environment and shutdown`, async function () {
        this.timeout(0);
        // Unregister the add-in
        await stopDebugging(manifestPath);
    }),
    describe(`Test Excel Desktop taskpane project`, function () {
        let testValues: any = [];
        it("Start test server", async function () {
            // Start test server and ping to ensure it's started
            testServer = new officeAddinTestServer.TestServer(testServerPort);
            const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
            const serverResponse = await officeAddinTestHelpers.pingTestServer(testServerPort);
            assert.strictEqual(testServerStarted, true);
            assert.strictEqual(serverResponse["status"], 200);
        });
        it("Validate expected result count", async function () {
            await startDebugging({manifestPath, appType: AppType.Desktop, app: toOfficeApp(hosts[0]), enableDebugging: false});
            this.timeout(0);
            testValues = await testServer.getTestResults();
            assert.strictEqual(testValues.length > 0, true);
        });
        it("Validate expected result name", async function () {
            assert.strictEqual(testValues[0].resultName, "fill-color");
        });
        it("Validate expected result", async function () {
            assert.strictEqual(testValues[0].resultValue, testValues[0].expectedValue);
        });
        it("Stop test server", async function () {
            this.timeout(0);
            const stopTestServer = await testServer.stopTestServer();
            assert.strictEqual(stopTestServer, true);
            testServer = undefined;
        });
    });
    describe(`Test Excel Web taskpane project`, function () {
        let testValues: any = [];
        it("Start test server", async function () {
            // Start test server and ping to ensure it's started
            testServer = new officeAddinTestServer.TestServer(testServerPort);
            const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
            const serverResponse = await officeAddinTestHelpers.pingTestServer(testServerPort);
            assert.strictEqual(testServerStarted, true);
            assert.strictEqual(serverResponse["status"], 200);
        });
        it("Validate expected result count", async function () {
            await startDebugging({manifestPath, appType: AppType.Web, app: toOfficeApp(hosts[0]), document: excelDoc});
            this.timeout(0);
            testValues = await testServer.getTestResults();
            assert.strictEqual(testValues.length > 0, true);
        });
        it("Validate expected result name", async function () {
            assert.strictEqual(testValues[0].resultName, "fill-color");
        });
        it("Validate expected result", async function () {
            assert.strictEqual(testValues[0].resultValue, testValues[0].expectedValue);
        });
        it("Stop test server", async function () {
            this.timeout(0);
            const stopTestServer = await testServer.stopTestServer();
            assert.strictEqual(stopTestServer, true);
            testServer = undefined;
        });
    });
    describe(`Test Word Desktop taskpane project`, function () {
        let testValues: any = [];
        it("Start test server", async function () {
            // Start test server and ping to ensure it's started
            testServer = new officeAddinTestServer.TestServer(testServerPort);
            const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
            const serverResponse = await officeAddinTestHelpers.pingTestServer(testServerPort);
            assert.strictEqual(testServerStarted, true);
            assert.strictEqual(serverResponse["status"], 200);
        });
        it("Validate expected result count", async function () {
            await startDebugging({manifestPath, appType: AppType.Desktop, app: toOfficeApp(hosts[1]), enableDebugging: false});
            this.timeout(0);
            testValues = await testServer.getTestResults();
            assert.strictEqual(testValues.length > 0, true);
        });
        it("Validate expected result name", async function () {
            assert.strictEqual(testValues[0].resultName, "output-message");
        });
        it("Validate expected result", async function () {
            assert.strictEqual(testValues[0].resultValue, testValues[0].expectedValue);
        });
        it("Stop test server", async function () {
            this.timeout(0);
            const stopTestServer = await testServer.stopTestServer();
            assert.strictEqual(stopTestServer, true);
            testServer = undefined;

            const applicationClosed = await testHelpers.closeDesktopApplication(hosts[1]);
            assert.strictEqual(applicationClosed, true);
        });
    });
    describe(`Test Word Web taskpane project`, function () {
        let testValues: any = [];
        it("Start test server", async function () {
            // Start test server and ping to ensure it's started
            testServer = new officeAddinTestServer.TestServer(testServerPort);
            const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
            const serverResponse = await officeAddinTestHelpers.pingTestServer(testServerPort);
            assert.strictEqual(testServerStarted, true);
            assert.strictEqual(serverResponse["status"], 200);
        });
        it("Validate expected result count", async function () {
            await startDebugging({manifestPath, appType: AppType.Web, app: toOfficeApp(hosts[1]), document: wordDoc});
            this.timeout(0);
            testValues = await testServer.getTestResults();
            assert.strictEqual(testValues.length > 0, true);
        });
        it("Validate expected result name", async function () {
            assert.strictEqual(testValues[0].resultName, "output-message");
        });
        it("Validate expected result", async function () {
            assert.strictEqual(testValues[0].resultValue, testValues[0].expectedValue);
        });
        it("Stop test server", async function () {
            this.timeout(0);
            const stopTestServer = await testServer.stopTestServer();
            assert.strictEqual(stopTestServer, true);
            testServer = undefined;
        });
    });
});