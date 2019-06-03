import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { run } from "../../src/taskpane/powerpoint";
import * as testHelpers from "./test-helpers";
const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
    if (info.host === Office.HostType.PowerPoint) {
        const testServerResponse: object = await pingTestServer(port);
        if (testServerResponse["status"] == 200) {
            await runTest();
        }
    }
});

export async function runTest() {
    try {
        // Execute taskpane code
        await run();
        await testHelpers.sleep(2000);

        // Get output of executed taskpane code
        const result = await getSelectedData();

        // Send test results to test server
        testHelpers.addTestResult(testValues, "output-message", result, " Hello World!"); 
        await sendTestResults(testValues, port);
        testValues.pop();
    } catch (err) {
        throw new Error(`runTest() failed with error ${err}`);
    }
}

async function getSelectedData(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text, async (asyncResult) => {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error.message);
                }
                else { resolve(asyncResult.value); }
            });
    });
}