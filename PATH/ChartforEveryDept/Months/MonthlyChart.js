

            

          const datam1 = [
                { t: new Date("2022-6-1"), y: 18 },
                { t: new Date("2022-6-7"), y: 10 },
                { t: new Date("2022-6-8"), y: 9 },
                { t: new Date("2022-6-9"), y: 4 },
                { t: new Date("2022-6-10"), y: 3 }
            
            
            
            ]; 
          

 


            function createLabels(limit) {
                const times = datam1.map(o => o.t.getTime());
                const startTime = times[0];
                const endTime = times[times.length - 1];
                const tickGap = (endTime - startTime) / (limit - 1);
                const labels = [startTime];
                for (let i = 1; i < limit - 1; i++) {
                    labels.push(startTime + i * tickGap);
                }
                labels.push(endTime);
                return labels;
            }


            var myChart = new Chart(document.getElementById("examChart1"), {
                type: 'line',
                data: {
                    labels: createLabels(5),
                    datasets: [{
                        label: 'Production (Main)',
                        lineTension: 0,
                        fill: false,
                        data: datam1,
                        backgroundColor: [
                            '#87CEEB'],
                        borderColor: [
                            '#87CEEB'],
                        borderWidth: 1
                    }  

                    ]


                },
                options: {
                    responsive: true,
                    scales: {
                        xAxes: [{
                            type: 'time',
                            ticks: {
                                source: 'labels',
                                minRotation: 45
                            },
                            time: {
                                unit: 'day',
                                displayFormats: {
                                    day: 'MM/DD/YYYY'
                                },
                                tooltipFormat: 'MM/DD/YYYY'
                            }
                        }],
                        yAxes: [{
                            ticks: {
                                min: 0,
                                max: 100,
                                callback: function (value, index, values) { return value + '%'; }

                            },
                            gridLines: {
                                zeroLineColor: "rgba(0,255,0,1)"
                            },
                            scaleLabel: {
                                display: true,
                                labelString: 'Percentage of Absent People'
                            }
                        }]
                    }
                }
            });



        export{datam1}; 
        