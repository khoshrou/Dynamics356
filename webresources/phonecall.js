// get contact
async function xpd_montagu_get_contact(id) {
    try {
        const result = await Xrm.WebApi.retrieveRecord("contact", id);
        console.log('*** get contact: ', result);
        return result;
    } catch (error) {
        console.log('*** Get Contact: Exception', error);
    }
}

// when adding a phone call if call to is available, populate the phone number
async function SetPhoneNumber(executionContext) {
    console.log('SetPhoneNumber Starting ....');
    const formContext = executionContext.getFormContext();

    const list = formContext.getAttribute("to").getValue();
    console.log('*** Call To Value', list);
    if (list && list.length > 0) {
        const contact = await xpd_montagu_get_contact(list[0].id);
        if (contact) {
            console.log('Conatact Phone Numbers: ', contact.mobilephone, contact.telephone1, contact.telephone2, contact.telephone3);
            if (contact.mobilephone)
                formContext.getAttribute("phonenumber").setValue(contact.mobilephone);
            else if (contact.telephone1)
                formContext.getAttribute("phonenumber").setValue(contact.telephone1);
            else if (contact.telephone2)
                formContext.getAttribute("phonenumber").setValue(contact.telephone2);
            else if (contact.telephone3)
                formContext.getAttribute("phonenumber").setValue(contact.telephone3);
        }
    }
}