/**Access token required to retrive shifts and Tasks graph api details. */
/* Id token required to authroize API controller */
let accessToken;
let idToken;
function hideProfileAndError() {
    $("#login").hide();
    $("#content").hide();
}

function successfulLogin() {
    $("#login").hide();
    $("#loading").show();
    $("#content").show();
    microsoftTeams.getContext(function (context) {
        getUserInfo(context.userPrincipalName);
        getShiftDetails(context.userObjectId);
       loadTeamMembers(context.userPrincipalName);
    });
    getTeamsConfiguration();
    loadNewsData();
    //loadAnnoucement();
}

function enableLogin() {
    $("#login").show();
    $("#loading").hide();
    $("#content").hide();
}

$(document).ready(function () {
    microsoftTeams.initialize();
    $(".horizontal .progress-fill span").each(function () {
        let percent = $(this).html();
        $(this).parent().css("width", percent);
    });
    $(".vertical .progress-fill span").each(function () {
        let percent = $(this).html();
        let pTop = 100 - percent.slice(0, percent.length - 1) + "%";
        $(this).parent().css({
            height: percent,
            top: pTop,
        });
    });
    enableLogin();
    $(document).ajaxStop(function () {
        $('#loading').hide();
    });
});

function getUserInfo(principalName) {
    if (principalName) {
        let graphUrl;
        graphUrl = "https://graph.microsoft.com/v1.0/users/" + principalName;
        $.ajax({
            url: graphUrl,
            type: "GET",
            beforeSend: function (request) {
                request.setRequestHeader("Authorization", "Bearer " + accessToken);
            },
            success: function (profile) {
                let name = profile.displayName;
                let userNameArray = name.split(' ');
                let myDate = new Date();
                let hrs = myDate.getHours();
                let greet;
                if (hrs < 12) {
                    greet = 'Bom dia, ';
                    $('#banner').css('background-image', "url('../images/gudmorning.png')");
                }
                else if (hrs >= 12 && hrs <= 17) {
                    greet = 'Boa tarde, ';
                    $('#banner').css('background-image', "url('../images/gudafternoon.png')");
                }
                else if (hrs >= 17 && hrs <= 24) {
                    greet = 'Boa noite, ';
                    $('#banner').css('background-image', "url('../images/gudevening.png')");
                }
                $('#greet').text(greet + userNameArray[0] + '!');
            },
            error: function () {
                console.log("Failed");
            },
            complete: function (data) {
            }
        });
    }
}

function getShiftDetails(objectId) {
    if (objectId) {
        let dateTimeNow = new Date().toISOString();
        let shiftFromDate = new Date();
        shiftFromDate.setDate(shiftFromDate.getDate() - 1);
        let graphShiftsUrl = "https://graph.microsoft.com/beta/teams/" + teamId + "/schedule/shifts?$filter=sharedShift/startDateTime ge " + shiftFromDate.toISOString();
        let graphTemp = [];
        do {
            $.ajax({
                url: graphShiftsUrl,
                type: "GET",
                async: false,
                beforeSend: function (request) {
                    request.setRequestHeader("Authorization", "Bearer " + accessToken);
                },
                success: function (response) {
                    if (response !== null) {
                        graphShiftsUrl = response["@odata.nextLink"];
                        graphTemp = graphTemp.concat(response.value);
                    } else {
                        console.log("Something went wrong");
                    }
                },
                error: function () {
                    console.log("Failed");
                }
            });
        }
        while (graphShiftsUrl)
        graphTemp.sort(sortShifts);
        let shift = graphTemp.find(s => (s.userId === objectId) && ((s.sharedShift.startDateTime <= dateTimeNow && s.sharedShift.endDateTime >= dateTimeNow) || s.sharedShift.startDateTime >= dateTimeNow));
        if (shift) {
            if (shift.sharedShift.startDateTime >= dateTimeNow) {
                $('#tasksCount').text('Seu próximo turno é:');
                getTaskDetails();
                $('#tasks').show();
                $('#survey').show();
            }
            else {
                getTaskDetails();
                $('#tasks').show();
                $('#survey').show();
                $('#tasksCount').text('Aproveite o seu turno e reveja as tarefas atribuídas a você.');
            }
            setShiftCard(shift);
        }
        else {
            $('#shiftHours').text('Não há turnos disponíveis.');
            $('#tasks').show();
            $('#survey').show();
        }
    }
}

function setShiftCard(item) {
    $('#shiftName').text(item.sharedShift.displayName);
    $('#shiftHours').text(new Date(item.sharedShift.startDateTime).toLocaleTimeString(navigator.language, { hour: '2-digit', minute: '2-digit' }) + ' - ' + new Date(item.sharedShift.endDateTime).toLocaleTimeString(navigator.language, { hour: '2-digit', minute: '2-digit' }));
    $('#shiftDate').text(new Date(item.sharedShift.startDateTime).getDate());
    $('#shiftDay').text(new Date(item.sharedShift.startDateTime).toString().split(' ')[0]);
    if (item.sharedShift.theme.includes('dark')) {
        $('#line').css('background', item.sharedShift.theme.substr(4).toLowerCase());
    }
    else {
        $('#line').css('background', item.sharedShift.theme);
    }
}

function getTaskDetails() {
    let graphTaskUrl = "https://graph.microsoft.com/v1.0/me/planner/tasks";
    $.ajax({
        url: graphTaskUrl,
        type: "GET",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", "Bearer " + accessToken);
        },
        success: function (response) {
            if (response !== null) {
                let arr = response.value;
                arr.sort(sortTasks);
                let taskUrl = "https://teams.microsoft.com/l/entity/" + tasksAppId + "/teamstasks.personalApp.mytasks?webUrl=https%3A%2F%2Fretailservices.teams.microsoft.com%2Fui%2Ftasks%2FpersonalApp%2Falltasklists&context=%7B%22subEntityId%22%3A%22%2FtaskListType%2FsmartList%2FSL_AssignedToMe%2Fplan%2F";
                let counter = 0;
                $.each(arr, function (i, item) {
                    if (item.completedDateTime === null) {
                        $('#taskSubject' + counter).text(item.title);
                        $('#taskDueDate' + counter).text(new Date(item.dueDateTime).toLocaleDateString());
                        $('#taskDueDate' + counter).attr('onclick', "microsoftTeams.executeDeepLink('" + taskUrl + encodeURIComponent(item.planId) + encodeURIComponent('/task/') + encodeURIComponent(item.id) + encodeURIComponent('"}') + "');");
                        counter++;
                        if (counter === 3) {
                            $('#seemoretasks').show();
                            $('#seemoretasks').attr('onclick', "microsoftTeams.executeDeepLink('https://teams.microsoft.com/l/entity/" + tasksAppId + "/tasks');");
                            return false;
                        }
                    }
                });
                switch (counter) {
                    case 0:
                        $('#tasks').hide();
                        break;
                    case 1:
                        $('#task2, #task3').hide();
                        break;
                    case 2:
                        $('#task3').hide();
                }
            } else {
                console.log("Something went wrong");
            }
        },
        error: function () {
            console.log("Failed");
        }
    });
}

function getTeamsConfiguration() {
    $.ajax({
        type: "GET",
        url: "/TeamsConfig",
        contentType: "application/json; charset=utf-8",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", "Bearer " + idToken);
        },
        success: function (response) {
            if (response !== null) {
                $('#payStubs').attr('onclick', "microsoftTeams.executeDeepLink('" + response.deepLinkBaseUrl + response.payStubsAppId + "/5e5b262b-a205-4e8b-8f23-a05e1ea237fb');");
                $('#benefits').attr('onclick', "microsoftTeams.executeDeepLink('" + response.deepLinkBaseUrl + response.benefitsAppId + "/eb72c432-7789-467a-9867-aa1243023aa4');");
                $('#rewards').attr('onclick', "microsoftTeams.executeDeepLink('https://teams.microsoft.com/l/file/AA1A883E-3580-4E6C-9868-D48AD5E4B07A?tenantId=04e9115d-32bf-4fd8-879d-aeb09ec6ecfc&fileType=pdf&objectUrl=https%3A%2F%2Fm365valeria.sharepoint.com%2Fsites%2FEnergisa%2FDocumentos%20Compartilhados%2FGeneral%2FManuais%2FFolheto.pdf&baseUrl=https%3A%2F%2Fm365valeria.sharepoint.com%2Fsites%2FEnergisa&serviceName=teams&threadId=19:e844faad6db14701bfed0916def361bc@thread.tacv2&groupId=a9de951d-01e6-4f3f-964c-ba2c06f958ef');");                
                $('#kudos').attr('onclick', "microsoftTeams.executeDeepLink('https://teams.microsoft.com/l/channel/19%3abfa10d943a394f30977ed0957e2b7589%40thread.tacv2/Manuais?groupId=a9de951d-01e6-4f3f-964c-ba2c06f958ef&tenantId=04e9115d-32bf-4fd8-879d-aeb09ec6ecfc');");                
                $('#news,#newsLink1,#newsLink2,#newsLink3').attr('onclick', "microsoftTeams.executeDeepLink('" + response.deepLinkBaseUrl + response.newsAppId + "/news');");
                $('#shifts').attr('onclick', "microsoftTeams.executeDeepLink('" + response.deepLinkBaseUrl + response.shiftsAppId + "/schedule');");
                $('#survey').attr('onclick', "microsoftTeams.executeDeepLink('" + response.deepLinkBaseUrl + response.surveyAppId + "/surveys');");
                $('#report').attr('onclick', "microsoftTeams.executeDeepLink('" + response.deepLinkBaseUrl + response.reportAppId + "/report');");
                
            } else {
                console.log("Something went wrong");
            }
        },
        failure: function (response) {
            console.log(response.responseText);
        },
        error: function (response) {
            console.log(response.responseText);
        }
    });
}

function loadTeamMembers(mailId) {
    if (mailId) {
        $.ajax({
            url: "/TeamMemberDetails",
            type: "Get",
            beforeSend: function (request) {
                request.setRequestHeader("Authorization", "Bearer " + idToken);
            },
            async: false,
            success: function (response) {
                if (response !== null) {
                    let counter = 0;
                    let groupEmail = [];
                    let chatUrl = "https://teams.microsoft.com/l/chat/0/0?users=";
                    for (let i = 0; i < response.length; i++) {
                        if (response[i].userPrincipalName === mailId) {
                            response.splice(i, 1);
                        }
                    }
                    let newMembers = response.splice(0, 2);
                    $.each(response, function (i, item) {
                        $('#memberName' + counter).text(item.givenName);
                        //$('#memberPicture' + counter).attr('src', item.profilePhotoUrl);
                        groupEmail.push(item.userPrincipalName);
                        counter++;
                        if (counter === 5) {
                            return false;
                        }
                    });
                    $('#groupChat').attr('onclick', "microsoftTeams.executeDeepLink('" + chatUrl + groupEmail.toString() + "&topicName=" + encodeURIComponent("On-Shift Crew")+"&message=Hi');");
                    $.each(newMembers, function (i, item) {
                        //$('#newMemberName' + i).text(item.givenName);
                       // $('#newMemberDesignation' + i).text(item.jobTitle);
                        //$('#newMemberPicture' + i).attr('src', item.profilePhotoUrl);
                        $('#newMemberChat' + i).attr('onclick', "microsoftTeams.executeDeepLink('https://teams.microsoft.com/l/chat/0/0?users=28:9a4d51b4-7550-4601-99fe-9be768fac8cc');");
                   });
                }
            },
            failure: function (response) {
                console.log(response.responseText);
            },
            error: function (response) {
                console.log(response.responseText);
            }
        });
    }
}

function loadAnnoucement() {
    let card;
    $.ajax({
        type: "GET",
        url: "/AnnouncementAdaptiveCardDetails",
        contentType: "application/json; charset=utf-8",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", "Bearer " + idToken);
        },
        dataType: "json",
        success: function (response) {
            if (response !== null) {
                console.log(response);
                let data = response.value;
                let counter = 0;
                let rowKey;
                let channelUrl = 'https://teams.microsoft.com/l/channel/';
                $.each(data, function (i, item) {
                    if (item.partitionKey === 'SendingNotifications' && !!item.content && counter === 0) {
                        card = JSON.parse(item.content);
                        rowKey = item.rowKey;
                        counter++;
                    }
                    if (item.partitionKey === 'SentNotifications' && !!item.teamsInString && item.rowKey === rowKey) {
                        let teamsId = item.teamsInString;
                        teamsId = teamsId.replace('["', "");
                        teamsId = teamsId.replace('"]', "");
                        //$('#annoucement').attr('onclick', "microsoftTeams.executeDeepLink('" + channelUrl + encodeURIComponent(teamsId) + "/General?groupId=7efb60ac-63c6-46e2-8645-8ea283dbfd61&tenantId=c80f38d3-c04c-49bf-a48b-9d99278d4ac6');");
                        $('#annoucement').attr('onclick', "microsoftTeams.executeDeepLink('https://teams.microsoft.com/l/channel/19%3ae844faad6db14701bfed0916def361bc%40thread.tacv2/Geral?groupId=a9de951d-01e6-4f3f-964c-ba2c06f958ef&tenantId=04e9115d-32bf-4fd8-879d-aeb09ec6ecfc');");
                        counter++;
                    }
                    if (counter === 2) {
                        return false;
                    }
                });
            }
        },
        failure: function (response) {
            console.log(response.responseText);
        },
        error: function (response) {
            console.log('Error occured while getting news data' + response.responseText);
        },
        complete: function (response) {
            if (card) {
                let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
                    fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
                });
                adaptiveCard.parse(card);
                let renderedCard = adaptiveCard.render();
                $('#annoucement').append(renderedCard);
            }
        }
    });
}

function loadBannersData() {
    $.ajax({
        type: "GET",
        url: "/BannersData",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", "Bearer " + idToken);
        },
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            console.log(response);
            $('#bannersTitle1').text(response.valuebanners[0].title);
            $('#bannersDescription1').text(response.valuebanners[0].description);
           // $("#bannersImage1").attr("src", response.valuebanners[0].image);
            $('#bannersTitle2').text(response.valuebanners[1].title);
            $('#bannersDescription2').text(response.valuebanners[1].description);
            //$("#bannersImage2").attr("src", response.valuebanners[1].image);
            $('#bannersTitle3').text(response.valuebanners[2].title);
            $('#bannersDescription3').text(response.valuebanners[2].description);
            //$("#bannersImage3").attr("src", response.valuebanners[2].image);

        },
        failure: function (response) {
            console.log(response.responseText);
        },
        error: function (response) {
            console.log('Error occured while getting banners data' + response.responseText);
        }
    });
}

function loadNewsData() {
    $.ajax({
        type: "GET",
        url: "/NewsData",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", "Bearer " + idToken);
        },
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            console.log(response);
            $('#newsTitle1').text(response.value[0].title);
            $('#newsDescription1').text(response.value[0].description);
            // $("#newsImage1").attr("src", response.value[0].image);            
            $('#newsTitle2').text(response.value[1].title);
            $('#newsDescription2').text(response.value[1].description);
            //$("#newsImage2").attr("src", response.value[1].image);            
            $('#newsTitle3').text(response.value[2].title);
            $('#newsDescription3').text(response.value[2].description);
            //$("#newsImage3").attr("src", response.value[2].image);

        },
        failure: function (response) {
            console.log(response.responseText);
        },
        error: function (response) {
            console.log('Error occured while getting news data' + response.responseText);
        }
    });
}

function sortShifts(a, b) {
    let dateA = new Date(a.sharedShift.startDateTime);
    let dateB = new Date(b.sharedShift.startDateTime);
    return dateA > dateB ? 1 : -1;
};

function sortTasks(a, b) {
    let dateA = new Date(a.dueDateTime);
    let dateB = new Date(b.dueDateTime);
    return dateA > dateB ? 1 : -1;
};