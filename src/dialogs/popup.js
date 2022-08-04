(function () {
    "use strict";

    Office.onReady()
        .then(function () {

            // TODO1: Assign handler to the OK button.
            // console.log(document.getElementById("ok-button"))
            if (document.getElementById("ok-button")) {
                document.getElementById("ok-button").onclick = sendTypeReportToParentPage;
            }
            if (document.getElementById("productplant-button")) {
                document.getElementById("productplant-button").onclick = sendProductPlantToParentPage;
            }
            // console.log(document.getElementById("ok-button"))
            if (document.getElementById("worker-button")) {
                document.getElementById("worker-button").onclick = sendIdWorkerToParentPage;
            }

            if (document.getElementById("date-button")) {
                document.getElementById("date-button").onclick = sendTimeToParentPage;
            }

            if (document.getElementById("task-button")) {
                document.getElementById("task-button").onclick = sendTaskToParentPage;
            }

        });


    // document.getElementById("test").innerHTML = "Do Cong kien"
    // $("#test").text("Hello Jquery")
    if ($("#abcd").length > 0) {
        const urlParams = new URLSearchParams(location.search);
        for (const [key, value] of urlParams) {
            // console.log(`${key}:${value}`);
            if (key == "id") {
                getListUser(value)
                break
            }
        }
    }
    async function getListUser(value) {
        await fetch(`http://localhost:8080/productplant/worker/${value}`)
            .then((res) => res.json())
            .then(data => {
                //hiển thị data( danh sách các phân xưởng) ra list
                console.log(data)
                for (let i of data) {
                    // console.log($("#productPlantList"))
                    // console.log(i.name)
                    let element = `<li>
                        <input type="radio" name="workerName" id="workerItem-${i.id}" value="${i.id}">
                        <label for="workerItem-${i.id}">${i.name}</label>
                                </li>`
                    // console.log(element)
                    $("#workerList").append(element)
                }
                // console.log("1")
            });
    }

    if ($("#taskList").length > 0) {
        popupTask()
    }
    async function popupTask() {
        const urlParams = new URLSearchParams(location.search);
        
        for (const [key, value] of urlParams) {
            // console.log(`${key}:${value}`);
            if (key == "id") {
                await getProductPlantById(value)
                await getListTask(value)
                break
            }
        }
    }

    async function getListTask(value) {
        await fetch(`http://localhost:8080/productplant/task/${value}`)
            .then((res) => res.json())
            .then(data => {
                //hiển thị data( danh sách các phân xưởng) ra list
                console.log(data)
                for (let i of data) {
                    // console.log($("#productPlantList"))
                    // console.log(i.name)
                    let element = `<li>
                        <input type="radio" name="taskName" id="taskItem-${i.id}" value="${i.id}">
                        <label for="taskItem-${i.id}">${i.name}</label>
                                </li>`
                    $("#taskList").append(element)
                }
            });
    }

    async function getProductPlantById(value) {
        await fetch(`http://localhost:8080/productplant/${value}`)
            .then((res) => res.json())
            .then(data => {
                //hiển thị data( danh sách các phân xưởng) ra list
                console.log(data)
                $("#popup__task").text(data[0].name)
            });
    }

    




    // document.getElementById("abcd").innerHTML = localStorage.getItem("clientID")

    // TODO2: Create the OK button handler
    function sendTypeReportToParentPage() {
        const typeReport = document.querySelector('input[name="type_report"]:checked').value;
        Office.context.ui.messageParent(typeReport);
    }

    function sendProductPlantToParentPage() {
        const productPlantValue = document.querySelector('input[name="productPlantName"]:checked').value;
        Office.context.ui.messageParent(productPlantValue);
    }


    function sendIdWorkerToParentPage() {
        const id_worker_dialog = document.querySelector('input[name="workerName"]:checked').value;
        Office.context.ui.messageParent(id_worker_dialog);
    }


    function sendTimeToParentPage() {
        const timeValueDialog = document.getElementById("time-input").value;
        const timeValueDialog2 = document.getElementById("time-input2").value;
        Office.context.ui.messageParent(timeValueDialog + " " + timeValueDialog2);
    }

    function sendTaskToParentPage() {
        const id_task_dialog = document.querySelector('input[name="taskName"]:checked').value;
        Office.context.ui.messageParent(id_task_dialog);
    }
    




}());