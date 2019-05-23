import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
import * as testHelper from "office-addin-test-helpers";
import * as testServerInfra from "office-addin-test-server";
const manifestPath = path.resolve(`${process.cwd()}/test/manifest.xml`);
const port: number = 4201;
const testJsonFile: string = path.resolve(`${process.cwd()}/test/src/testData.json`);
const testJsonData = JSON.parse(fs.readFileSync(testJsonFile).toString());
const wefCache = path.join(process.env.USERPROFILE, `AppData/Local/Microsoft/Office/16.0/Wef`);

// Running on Windows only due to VSO.Bug 3377441: Office Addins fail to
// work on Mac if addin has previously been cached in the WEF folder
if (process.platform === "win32") {
    Object.keys(testJsonData.hosts).forEach(function (host) {
        const testServer = new testServerInfra.TestServer(port);
        const resultName = testJsonData.hosts[host].resultName;
        const resultValue: string = testJsonData.hosts[host].resultValue;
        let testValues: any = [];

        describe(`Test ${host} Task Pane Project`, function () {
            before("Test Server should be started", async function () {                
                const wefCacheCleared = await clearWefCache(wefCache);
                const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
                const serverResponse = await testHelper.pingTestServer(port);
                assert.equal(wefCacheCleared, true);
                assert.equal(testServerStarted, true);
                assert.equal(serverResponse["status"], 200);
            }),
                describe(`Start dev-server and sideload application: ${host}`, function () {
                    it(`Sideload should have completed for ${host} and dev-server should have started`, async function () {
                        this.timeout(0);
                        const startDevServer = await testHelper.startDevServer();
                        const sideloadApplication = await testHelper.sideloadDesktopApp(host, manifestPath);
                        assert.equal(startDevServer, true);
                        assert.equal(sideloadApplication, true);
                    });
                });
            describe(`Get test results for ${host} taskpane project`, function () {
                it("Validate expected result count", async function () {
                    this.timeout(0);
                    testValues = await testServer.getTestResults();
                    assert.equal(testValues.length > 0, true);
                });
                it("Validate expected result name", async function () {
                    assert.equal(testValues[0].Name, resultName);
                });
                it("Validate expected result", async function () {
                    assert.equal(testValues[0].Value, resultValue);
                });
            });
            after(`Teardown test environment and shutdown ${host}`, async function () {
                const stopTestServer = await testServer.stopTestServer();
                assert.equal(stopTestServer, true);
                const wefCacheCleared = await clearWefCache(wefCache);
                assert.equal(wefCacheCleared, true);
                const testEnvironmentTornDown = await testHelper.teardownTestEnvironment(host, host != 'Excel');
                assert.equal(testEnvironmentTornDown, true);
            });
        });
    });
}

async function clearWefCache(wefFolder: string): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        try {
            if (process.platform === "win32") {                 
                if (fs.existsSync(wefFolder)) {
                    fs.readdirSync(wefFolder).forEach(function (files, index) {
                        if (files.length) {
                            const curPath = path.join(wefFolder, files)

                            if (fs.lstatSync(curPath).isDirectory()) {
                                clearWefCache(curPath);
                            }
                            else {
                                fs.unlinkSync(curPath);
                            }
                        } else {
                            resolve(true);
                        }
                    });
                    if (wefFolder != wefCache) {
                        fs.rmdirSync(wefFolder);
                    }
                    else {
                        resolve(true);
                    }
                }
            }
        } catch (err) {
            return reject(false);
        }
    });
}


