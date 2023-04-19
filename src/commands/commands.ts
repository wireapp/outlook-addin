/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { addMeetingLink } from "../calendarIntegration/addMeetingLink";

export let mailboxItem;

// Office is ready. Init
Office.onReady(function () {
  mailboxItem = Office.context.mailbox.item;
});

// Register the functions.
Office.actions.associate("addMeetingLink", addMeetingLink);
