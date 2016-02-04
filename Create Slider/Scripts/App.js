'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage()
{
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        $('#siteContent').attr('href', decodeURIComponent(getQueryStringParameter('SPHostUrl')) + "/_layouts/15/viewlsts.aspx");
        $('#createList').click(createSliderList);
    });

    //#1 Create List
    function createSliderList() {
        var imgTarget = decodeURIComponent(getQueryStringParameter('SPHostUrl')) + '/SiteAssets/l_Sliderx96.png'

        // Create an announcement SharePoint list with the name that the user specifies.
        var hostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
        var currentcontext = new SP.ClientContext.get_current();
        var hostcontext = new SP.AppContextSite(currentcontext, hostUrl);
        var hostweb = hostcontext.get_web();

        //Set ListCreationInfomation()
        var listCreationInfo = new SP.ListCreationInformation();
        listCreationInfo.set_title('Sliderr');
        listCreationInfo.set_templateType(SP.ListTemplateType.genericList);
        var newList = hostweb.get_lists().add(listCreationInfo);
        newList.set_imageUrl(imgTarget);
        newList.update();

        //Set column data
        var newCols = [
            "<Field Type='Note' DisplayName='Body' Required='FALSE' EnforceUniqueValues='FALSE' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' StaticName='Body' Name='Body'/>",
            "<Field Type='URL' DisplayName='Background Image' Required='FALSE' EnforceUniqueValues='FALSE' Format='Hyperlink' StaticName='BackgroundImage' Name='BackgroundImage'/>",
            "<Field Type='Boolean' DisplayName='Enabled' EnforceUniqueValues='FALSE' StaticName='Enabled' Name='Enabled'><Default>1</Default></Field>",
            "<Field Type='DateTime' DisplayName='Expire' Required='FALSE' EnforceUniqueValues='FALSE' Format='DateTime' FriendlyDisplayFormat='Disabled' StaticName='Expire' Name='Expire'/>",
            "<Field Type='Number' DisplayName='Order' Required='FALSE' EnforceUniqueValues='FALSE' StaticName='Order0' Name='Order0'/>"
        ];
        var newListWithColumns;
        for (var i = 0; i < newCols.length; i++) {
            newListWithColumns = newList.get_fields().addFieldAsXml(newCols[i], true, SP.AddFieldOptions.defaultValue);
        }

        //final load/execute
        context.load(newListWithColumns);
        context.executeQueryAsync(function () {
            console.log('Slider list created successfully!');
            uploadTileImage()
        },
        function (sender, args) {
            console.error(sender);
            console.error(args);
            alert('Failed to create the Slider list. ' + args.get_message());
        });
    }

    //#2 Upload Tile Image
    function uploadTileImage() {
        BinaryUpload.Uploader().Upload("/images/l_Sliderx96.png", "/SiteAssets/l_Sliderx96.png");
        //setImageUrl();
        createConfigList();
    }

    //#3 Set Tile Image
    function setImageUrl() {
        var urlTarget = decodeURIComponent(getQueryStringParameter('SPAppWebUrl') + '/_api/SP.AppContextSite(@target)/web/lists/getByTitle(\'Sliderr\')?@target=\'' + getQueryStringParameter('SPHostUrl') + '\'');
        var imgTarget = decodeURIComponent(getQueryStringParameter('SPHostUrl')) + '/SiteAssets/l_Sliderx96.png'
        console.log(urlTarget);
        $.ajax({
            method: 'POST',
            url: urlTarget,
            data: JSON.stringify({ '__metadata': { 'type': 'SP.List' }, 'ImageUrl': imgTarget }),
            contentType: 'application/json;odata=verbose',
            headers: {
                'accept': 'application/json;odata=verbose',
                'X-RequestDigest': $('#__REQUESTDIGEST').val(),
                'X-HTTP-Method': 'MERGE',
                'If-Match': '*'
            },
            success: function (data) {
                console.log('successfully set ImageUrl');
                $('#sliderListContainer').show();
                alert('Slider list created successfully! Click on the Slider List to continue.');
            },
            error: function (err) {
                console.error(err);
                alert('Failed to set List.ImageUrl. Please notifiy your SharePoint Admin.');
            }
        });
    }

    //#4 create spConfig List
    function createConfigList() {
        var hostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
        var currentcontext = new SP.ClientContext.get_current();
        var hostcontext = new SP.AppContextSite(currentcontext, hostUrl);
        var hostweb = hostcontext.get_web();

        //Set ListCreationInfomation()
        var listCreationInfo = new SP.ListCreationInformation();
        listCreationInfo.set_title('spConfig');
        listCreationInfo.set_templateType(SP.ListTemplateType.genericList);
        var newList = hostweb.get_lists().add(listCreationInfo);
        newList.set_hidden(true);
        newList.set_onQuickLaunch(false);
        newList.update();

        //Set column data
        var newListWithColumns = newList.get_fields().addFieldAsXml("<Field Type='Note' DisplayName='Value' Required='FALSE' EnforceUniqueValues='FALSE' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' StaticName='Value' Name='Value'/>", true, SP.AddFieldOptions.defaultValue);

        //final load/execute
        context.load(newListWithColumns);
        context.executeQueryAsync(function () {
            console.log('spConfig list created successfully!');
        },
        function (sender, args) {
            console.error(sender);
            console.error(args);
            alert('Failed to create the spConfig list. ' + args.get_message());
        });
    }

}

function getQueryStringParameter(param) {
    var params = document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == param) {
            return singleParam[1];
        }
    }
}
