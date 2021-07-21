// Add Remove To Outlook

//refresh the form

function refreshForm() {
    Xrm.Utility.openEntityForm(Xrm.Page.data.entity.getEntityName(), Xrm.Page.data.entity.getId());
}

async function alertDialog({ title, subtitle, text, confirmButtonLabel, cancelButtonLabel }) {
    try {
        const alertStrings = {
            confirmButtonLabel: confirmButtonLabel || "Yes", cancelButtonLabel: cancelButtonLabel || 'No', text, title: title || 'Message', subtitle: subtitle || ''
        };
        const alertOptions = { height: 120, width: 260 };
        const result = await Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
        console.log('Result', result);
        return true;
    } catch (error) {
        console.log('alertDialog:Exception ', error);
    }
    return false;
}

async function confirmDialog({ title, subtitle, text, confirmButtonLabel, cancelButtonLabel }) {
    try {
        const confirmStrings = {
            confirmButtonLabel: confirmButtonLabel || "Yes", cancelButtonLabel: cancelButtonLabel || 'No', text, title: title || 'Confirm', subtitle: subtitle || ''
        };
        const confirmOptions = { height: 120, width: 450 };
        const result = await Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions);
        console.log('Result', result);
        return !!(result && result.confirmed);
    } catch (error) {
        console.log('confirmDialog:Exception ', error);
    }
    return false;
}

// get user id and entity id
function getIds() {
    const entityId = Xrm.Utility.getPageContext().input.entityId;
    console.log('**** Entity Id', entityId);
    const userSettings = Xrm.Utility.getGlobalContext().userSettings;
    const uid = userSettings.userId;
    console.log('*** UserId = ' + uid);
    return { uid, entityId };
}

// get existing relations
async function xpd_montagu_get_relations() {
    try {
        const { uid, entityId } = getIds();

        const fetchXml = `?fetchXml=
        <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
        <entity name="ts_montagurelationship">
            <attribute name="ts_montagurelationshipid" />
            <attribute name="ts_name" />
            <attribute name="createdon" />
            <order attribute="ts_name" descending="false" />
            <filter type="and">
            <condition attribute="ts_relationship" operator="eq" value="717750000" />
            </filter>
            <link-entity name="systemuser" from="systemuserid" to="ts_userid" link-type="inner" alias="ad">
            <filter type="and">
                <condition attribute="systemuserid" operator="eq" uitype="systemuser" value="${uid}" />
            </filter>
            </link-entity>
            <link-entity name="contact" from="contactid" to="ts_contact" link-type="inner" alias="ae">
            <filter type="and">
                <condition attribute="contactid" operator="eq" uitype="contact" value="${entityId}" />
            </filter>
            </link-entity>
        </entity>
        </fetch>
        `;
        const result = await Xrm.WebApi.retrieveMultipleRecords("ts_montagurelationship", fetchXml)
        console.log('*** Fetch xml Result Count: ', result.entities.length);

        return result.entities;
    } catch (error) {
        console.log('*** Get Relations: Exception', error);
    }

    return false;
}

// check if a relation exists
async function xpd_montagu_check_relation() {
    try {
        const entities = await xpd_montagu_get_relations();

        // enable if no data found
        return entities.length > 0 ? true : false;
    } catch (error) {
        console.log('*** Checking Relation Exists: Exception', error);
    }

    return false;
}

// check if add to outlook is enabled
async function xpd_montagu_check_addOutlook() {
    console.log('*** Checking add to outlook button');
    try {

        const exists = await xpd_montagu_check_relation();

        // enable if no data found
        return !exists;

    } catch (error) {
        console.log('*** Checking add to outlook button: Exception', error);
    }

    return false;
}

// run add to outlook 
async function xpd_montagu_run_addOutlook(formContext) {
    console.log('Add To Outlook Clicked', formContext ? 'OK' : 'NOK');
    try {
        const exists = await xpd_montagu_check_relation();
        if (!exists) {
            if (await confirmDialog({ text: 'Do you want to add the  selected contact to outlook?' })) {
                const { uid, entityId } = getIds();
                // remove braces
                const uid2 = uid.replace('{', '').replace('}', '');
                const entityId2 = entityId.replace('{', '').replace('}', '');
                const data = {
                    ts_relationship: 717750000,
                    "ts_UserId@odata.bind": `/systemusers(${uid2})`,
                    "ts_Contact@odata.bind": `/contacts(${entityId2})`
                };

                console.log('Create Data', data);

                Xrm.Utility.showProgressIndicator('Adding To Outlook');
                await Xrm.WebApi.createRecord("ts_montagurelationship", data);
                Xrm.Utility.closeProgressIndicator();
                refreshForm();
                //alertDialog({ text: 'Successfully added to outlook.' });
            }
        } else {
            alertDialog({ text: 'Already added.' });
        }
    } catch (error) {
        alertDialog({ text: 'Exception on Add: ' + error.message });
    }

}

// check if remove from outlook is enabled
async function xpd_montagu_check_removeOutlook(context) {
    console.log('*** Checking remove from outlook button');
    try {

        const exists = await xpd_montagu_check_relation();

        // enable if no data found
        return exists;

    } catch (error) {
        console.log('*** Checking add to outlook button: Exception', error);
    }

    return false;
}

// run add to outlook 
async function xpd_montagu_run_removeOutlook(context) {
    console.log('Remove From Outlook Clicked');
    try {
        const entities = await xpd_montagu_get_relations();
        if (entities && entities.length > 0) {
            if (await confirmDialog({ text: 'Do you want to remove the selected contct from outlook?' })) {
                Xrm.Utility.showProgressIndicator('Removing from Outlook');
                for (const entity of entities) {
                    console.log('Entity To Delete', entity);
                    const entityId = entity.ts_montagurelationshipid;
                    console.log('Entity Id', entityId);
                    await Xrm.WebApi.deleteRecord("ts_montagurelationship", entityId);
                }
                Xrm.Utility.closeProgressIndicator();
                refreshForm();
                //alert('Successfully removed from outlook.');
            }
        } else {
            alertDialog({ text: 'Already removed.' });
        }
    } catch (error) {
        alertDialog({ text: 'Exception on Remove: ' + error.message });
    }
}