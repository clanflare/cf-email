// Import necessary modules
import { Buffer } from "node:buffer";
globalThis.Buffer = Buffer;

import { EmailMessage } from "cloudflare:email";
import { createMimeMessage } from "mimetext";
import PostalMime from "postal-mime";

/**
 * Constants and Configuration
 */

// Discord API version and base URL
const API_VERSION = 10;
const DISCORD_API_URL = `https://discord.com/api/v${API_VERSION}`;

// Environment variables (to be initialized in the main function)
let GUILD_ID, TOKEN, ROLES_REQUIRED, CHANNEL_MAP, ATTACHMENTS_CHANNEL, LOG_CHANNEL_ID;

// Headers and request options
const myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");

// Rate limiting configuration
const MAX_RETRIES = 3;
const DEFAULT_RETRY_AFTER_MS = 1000; // 1 second

/**
 * Logging Variables
 */

// Array to collect log entries
let logEntries = [];

/**
 * Utility Functions
 */

// Convert a stream to an ArrayBuffer
async function streamToArrayBuffer(stream) {
  try {
    const arrayBuffer = await new Response(stream).arrayBuffer();
    return new Uint8Array(arrayBuffer);
  } catch (error) {
    console.error("Error converting stream to ArrayBuffer:", error);
    await logToDiscord("ERROR", `Error converting stream to ArrayBuffer: ${error.message}`);
    throw new Error("Failed to process email stream.");
  }
}

// Fetch with rate limit handling
async function fetchWithRateLimit(url, options, retries = MAX_RETRIES) {
  let response = await fetch(url, options);

  while (response.status === 429 && retries > 0) {
    const retryAfter = response.headers.get("Retry-After");
    const retryAfterMs = retryAfter
      ? parseFloat(retryAfter) * 1000
      : DEFAULT_RETRY_AFTER_MS;
    console.warn(`Rate limited. Retrying after ${retryAfterMs}ms`);
    await new Promise((resolve) => setTimeout(resolve, retryAfterMs));

    response = await fetch(url, options);
    retries--;
  }

  if (response.status === 429) {
    const errorMsg = "Rate limit exceeded, maximum retries reached.";
    console.error(errorMsg);
    await logToDiscord("ERROR", errorMsg);
    throw new Error(errorMsg);
  }

  return response;
}

// Truncate text safely
function truncateText(text, maxLength) {
  return text.length > maxLength ? `${text.substring(0, maxLength - 3)}...` : text;
}

// Decode HTML entities
function decodeHtmlEntities(text) {
  const entities = {
    nbsp: " ",
    amp: "&",
    lt: "<",
    gt: ">",
    quot: '"',
    apos: "'",
    "#39": "'",
  };
  return text.replace(/&([^;]+);/g, (match, entity) => entities[entity] || match);
}

// Extract text from HTML content
function extractTextFromHtml(htmlContent) {
  try {
    let text = htmlContent
      .replace(/<script[^>]*>([\s\S]*?)<\/script>/gi, "")
      .replace(/<style[^>]*>([\s\S]*?)<\/style>/gi, "")
      .replace(/<(br|\/p|p)[^>]*>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return decodeHtmlEntities(text);
  } catch (error) {
    console.error("Error extracting text from HTML:", error);
    return "(Error extracting text content)";
  }
}

/**
 * Logging Functions
 */

// Add a log entry to the logEntries array
function logToDiscord(level, message) {
  const timestamp = new Date().toISOString();
  logEntries.push({ timestamp, level, message });
}

/**
 * Email Handling Functions
 */

// Parse the email using PostalMime
async function parseEmail(event) {
  try {
    logToDiscord("INFO", "Starting to parse email.");
    const rawEmail = await streamToArrayBuffer(event.raw);
    const parser = new PostalMime();
    const parsedEmail = await parser.parse(rawEmail);
    logToDiscord("INFO", `Email parsed successfully.\nSubject: ${parsedEmail.subject || "No Subject"}`);
    return parsedEmail;
  } catch (error) {
    console.error("Error parsing email:", error);
    logToDiscord("ERROR", `Error parsing email: ${error.message}`);
    throw new Error("Failed to parse email content.");
  }
}

// Sanitize data to avoid logging sensitive information
function sanitizeData(data) {
  // Implement sanitization logic as needed
  // For example, redact email addresses or other sensitive fields
  return data;
}

// Send an auto-reply email
async function sendAutoReply(event, parsedEmail, errorMsg = null) {
  try {
    const timestamp = new Date().toISOString();
    const subject = `Re: ${parsedEmail.subject || "No Subject"}`;
    const originalRecipient = parsedEmail.to[0]?.address.split("@")[0] || "Unknown";
    const messageData = errorMsg
      ? `
        <p>Dear Sender,</p>
        <p>An error occurred while processing your email. Please review the error details below:</p>
        <blockquote style="border-left: 2px solid #ccc; padding-left: 10px; color: #555;">
          <p><strong>Error Details:</strong> ${errorMsg}</p>
        </blockquote>
        <p>If you need further assistance, feel free to contact support.</p>
        <p>Best regards,<br/>Clanflare</p>
        <hr />
        <p style="font-size: 0.9em; color: #888;">Timestamp: ${timestamp}</p>
      `
      : `
        <p>Dear Sender,</p>
        <p>Thank you for your email. This is an automated reply to confirm that your email with the subject "<strong>${parsedEmail.subject || "No Subject"
      }</strong>" has been successfully delivered to the recipient: <strong>${originalRecipient}</strong>.</p>
        <p>If you have any further inquiries, please feel free to reach out.</p>
        <p>Best regards,<br/>Clanflare</p>
        <hr />
        <p style="font-size: 0.9em; color: #888;">Timestamp: ${timestamp}</p>
      `;

    // Fallback to event.from if parsedEmail.from is missing
    const fromAddress = parsedEmail.from?.[0]?.address || parsedEmail.from?.address || event.from;
    const toAddress = parsedEmail.to?.[0]?.address || event.to;

    if (!fromAddress) {
      const warningMsg = "Original email is not repliable: Missing 'from' address.";
      console.warn(warningMsg);
      logToDiscord("WARN", warningMsg);
      return;
    }

    const msg = createMimeMessage();
    msg.setSender({ name: "Auto-replier", addr: toAddress });
    msg.setRecipient(fromAddress);
    msg.setSubject(subject);

    if (parsedEmail.messageId) {
      msg.setHeader("In-Reply-To", parsedEmail.messageId);
    }

    msg.addMessage({
      contentType: "text/html",
      data: messageData,
    });

    const message = new EmailMessage(toAddress, fromAddress, msg.asRaw());

    // Attempt to reply to the original email
    try {
      await event.reply(message);
      logToDiscord("INFO", "Auto-reply sent successfully using event.reply().");
    } catch (replyError) {
      console.warn("event.reply() failed, attempting to send a new email:", replyError);
      logToDiscord("WARN", `event.reply() failed: ${replyError.message}. Attempting to send a new email.`);
      // If event.reply() fails, send a new email instead
      await event.send(message);
      logToDiscord("INFO", "Auto-reply sent using event.send().");
    }
  } catch (error) {
    console.error("Error sending auto-reply:", error);
    logToDiscord("ERROR", `Error sending auto-reply: ${error.message}`);
  }
}

/**
 * Discord API Interaction Functions
 */

// Fetch guild data from Discord API
async function fetchGuildData() {
  try {
    const url = `${DISCORD_API_URL}/guilds/${GUILD_ID}`;
    const response = await fetchWithRateLimit(url, {
      method: "GET",
      headers: myHeaders,
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch guild data. Status: ${response.status}`);
    }

    const guildData = await response.json();
    logToDiscord("INFO", `Fetched guild data: ${guildData.name}`);
    return guildData;
  } catch (error) {
    console.error("Error fetching guild data:", error);
    logToDiscord("ERROR", `Error fetching guild data: ${error.message}`);
    throw new Error("Failed to fetch guild information.");
  }
}

// Get the guild icon URL
function getGuildIconURL(guildData) {
  if (guildData.icon) {
    const fileExtension = guildData.icon.startsWith("a_") ? ".gif" : ".png";
    return `https://cdn.discordapp.com/icons/${guildData.id}/${guildData.icon}${fileExtension}`;
  }
  return null;
}

// Find Discord member by username
async function findDiscordMember(username) {
  try {
    const fetchMemberURL = `${DISCORD_API_URL}/guilds/${GUILD_ID}/members/search?query=${encodeURIComponent(
      username
    )}&limit=1000`;

    const response = await fetchWithRateLimit(fetchMemberURL, {
      method: "GET",
      headers: myHeaders,
    });
    if (!response.ok) {
      throw new Error(`Failed to fetch Discord members. Status: ${response.status}`);
    }

    const members = await response.json();
    const member = members.find((mem) => mem.user.username === username);
    if (member) {
      logToDiscord("INFO", `Found Discord member: ${member.user.username}#${member.user.discriminator}`);
    } else {
      logToDiscord("WARN", `Discord member not found for username: ${username}`);
    }
    return member;
  } catch (error) {
    console.error("Error fetching Discord member:", error);
    logToDiscord("ERROR", `Error fetching Discord member: ${error.message}`);
    throw new Error(`Failed to find Discord member with username: ${username}`);
  }
}

// Check if the member has the required role(s)
function hasRequiredRoles(member) {
  if (!member) return false;
  const memberRoles = member.roles; // Array of role IDs the member has

  if (ROLES_REQUIRED.length === 0) return true; // No roles required

  const hasRoles = ROLES_REQUIRED.some((role) => memberRoles.includes(role));
  return hasRoles;
}

// Create a DM channel with the user
async function createDmChannel(recipientId) {
  try {
    const createDmURL = `${DISCORD_API_URL}/users/@me/channels`;
    const response = await fetchWithRateLimit(createDmURL, {
      method: "POST",
      body: JSON.stringify({ recipient_id: recipientId }),
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bot ${TOKEN}`,
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to create DM channel. Status: ${response.status}`);
    }

    const dmChannel = await response.json();
    logToDiscord("INFO", `Created DM channel with user ID: ${recipientId}`);
    return dmChannel;
  } catch (error) {
    console.error("Error creating DM channel:", error);
    logToDiscord("ERROR", `Error creating DM channel: ${error.message}`);
    throw new Error("Failed to create a direct message channel with the member.");
  }
}

// Send a batch of embeds to a Discord channel
async function sendEmbedBatch(channelId, payload, attachments = null) {
  const options = {
    method: "POST",
    headers: {
      Authorization: `Bot ${TOKEN}`,
    },
  };

  if (attachments) {
    // Sending with attachments
    const formData = new FormData();
    formData.append("payload_json", JSON.stringify(payload));
    attachments.forEach((attachment, index) => {
      formData.append(`files[${index}]`, attachment.blob, attachment.filename);
    });
    options.body = formData;
  } else {
    // Sending without attachments
    options.headers["Content-Type"] = "application/json";
    options.body = JSON.stringify(payload);
  }

  const response = await fetchWithRateLimit(
    `${DISCORD_API_URL}/channels/${channelId}/messages`,
    options
  );

  if (!response.ok) {
    const errorText = await response.text();
    logToDiscord("ERROR", `Failed to send embed batch: ${errorText}`);
    throw new Error(`Failed to send message. Status: ${response.status}. Error: ${errorText}`);
  }

  return await response.json();
}

// Send multiple embeds with batching
async function sendEmbedsWithBatching(channelId, embeds, attachments = null) {
  const MAX_EMBEDS_PER_MESSAGE = 10;
  const MAX_TOTAL_EMBED_SIZE = 6000; // Maximum total size of embeds per message
  const messages = [];

  let currentBatch = [];
  let currentTotalSize = 0;

  for (const embed of embeds) {
    const embedSize = JSON.stringify(embed).length;

    // Check if adding the embed would exceed limits
    if (
      currentTotalSize + embedSize > MAX_TOTAL_EMBED_SIZE ||
      currentBatch.length >= MAX_EMBEDS_PER_MESSAGE
    ) {
      // Send current batch
      const payload = { embeds: currentBatch };
      await sendEmbedBatch(channelId, payload);
      messages.push(payload);

      // Reset batch
      currentBatch = [];
      currentTotalSize = 0;
    }

    // Add embed to batch
    currentBatch.push(embed);
    currentTotalSize += embedSize;
  }

  // Send any remaining embeds
  if (currentBatch.length > 0) {
    const payload = { embeds: currentBatch };
    await sendEmbedBatch(channelId, payload, attachments);
    messages.push(payload);
  }

  logToDiscord("INFO", `Sent ${embeds.length} embeds to channel ID: ${channelId}`);
  return messages;
}

// Send the text content as an attachment
async function sendTextAttachment(channelId, textContent, attachmentLinks) {
  try {
    const formData = new FormData();

    // Combine text content and attachment links
    let fullTextContent = textContent;

    if (attachmentLinks && attachmentLinks.length > 0) {
      const attachmentsText = attachmentLinks
        .map((link, index) => `Attachment ${index + 1}: ${link}`)
        .join("\n");
      fullTextContent += `\n\nAttachments:\n${attachmentsText}`;
    }

    // Create a Blob from the content
    const blob = new Blob([fullTextContent], { type: "text/plain" });
    formData.append("files[0]", blob, "full_message.txt");

    const response = await fetchWithRateLimit(
      `${DISCORD_API_URL}/channels/${channelId}/messages`,
      {
        method: "POST",
        headers: {
          Authorization: `Bot ${TOKEN}`,
        },
        body: formData,
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      logToDiscord("ERROR", `Failed to send text attachment: ${errorText}`);
      throw new Error(
        `Failed to send message with attachment. Status: ${response.status}. Error: ${errorText}`
      );
    }

    logToDiscord("INFO", `Sent text attachment to channel ID: ${channelId}`);
    return await response.json();
  } catch (error) {
    console.error("Error sending text attachment:", error);
    logToDiscord("ERROR", `Error sending text attachment: ${error.message}`);
    throw new Error("Failed to send the text attachment.");
  }
}

// Upload attachments to a Discord channel and return their URLs
async function uploadAttachmentsToChannel(channelId, attachments) {
  try {
    const formData = new FormData();

    // Append attachments (up to Discord's limit per message)
    attachments.forEach((attachment, index) => {
      const blob = new Blob([attachment.content], { type: attachment.contentType });
      formData.append(`files[${index}]`, blob, attachment.filename);
    });

    const sendMessageURL = `${DISCORD_API_URL}/channels/${channelId}/messages`;

    const response = await fetchWithRateLimit(sendMessageURL, {
      method: "POST",
      headers: {
        Authorization: `Bot ${TOKEN}`,
      },
      body: formData,
    });

    if (!response.ok) {
      const errorText = await response.text();
      logToDiscord("ERROR", `Failed to upload attachments: ${errorText}`);
      throw new Error(
        `Failed to upload attachments. Status: ${response.status}. Error: ${errorText}`
      );
    }

    const messageData = await response.json();

    // Extract attachment URLs
    if (messageData.attachments && messageData.attachments.length > 0) {
      const urls = messageData.attachments.map((attachment) => attachment.url);
      logToDiscord("INFO", `Uploaded ${urls.length} attachments to channel ID: ${channelId}`);
      return urls;
    } else {
      const errorMsg = "No attachment URLs found in message data.";
      console.error(errorMsg);
      logToDiscord("WARN", errorMsg);
      return [];
    }
  } catch (error) {
    console.error("Error uploading attachments:", error);
    logToDiscord("ERROR", `Error uploading attachments: ${error.message}`);
    return [];
  }
}

// Handle attachments and return their links
async function handleAttachments(attachments) {
  const attachmentLinks = [];

  // Batch attachments according to Discord's limits
  const BATCH_SIZE = 10;
  const batches = [];

  for (let i = 0; i < attachments.length; i += BATCH_SIZE) {
    batches.push(attachments.slice(i, i + BATCH_SIZE));
  }

  for (const batch of batches) {
    const links = await uploadAttachmentsToChannel(ATTACHMENTS_CHANNEL, batch);
    if (links && links.length > 0) {
      attachmentLinks.push(...links);
    }
  }

  logToDiscord("INFO", `Handled ${attachments.length} attachments.`);
  return attachmentLinks;
}

// Send the email content as embeds to Discord
async function sendEmbedMessage(channelId, parsedEmail, event, attachmentLinks) {
  try {
    // Fetch guild data and icon
    const guildData = await fetchGuildData();
    const thumbnailUrl = getGuildIconURL(guildData);

    // Prepare email fields with fallback
    const fromField = parsedEmail.from?.address || parsedEmail.from?.[0]?.address || event.from || "Unknown";
    const toField = parsedEmail.to?.map((addr) => addr.address).join(", ") || event.to || "Unknown";

    // Extract email text content
    let emailTextContent = parsedEmail.text || extractTextFromHtml(parsedEmail.html);
    emailTextContent = emailTextContent || "(No text content)";

    // Sanitize email text content for privacy
    emailTextContent = sanitizeData(emailTextContent);

    // Maximum lengths
    const MAX_DESCRIPTION_LENGTH = 4096;
    const MAX_TOTAL_EMBED_SIZE = 6000; // Max total size of embeds per message

    // Construct the initial embed
    const title = `📧 ${parsedEmail.subject || "New Email Received"}`;
    const footerText = "📬 Sent via Clanflare Email System";
    const timestamp = new Date().toISOString();

    let embeds = [
      {
        title,
        description: truncateText(
          `**👤 From:** ${fromField}\n**📩 To:** ${toField}\n**📅 Date:** ${parsedEmail.date ||
          timestamp}\n**📎 Attachments:** ${attachmentLinks.length}`,
          MAX_DESCRIPTION_LENGTH
        ),
        footer: { text: footerText },
        timestamp,
        thumbnail: { url: thumbnailUrl },
      },
    ];

    // Add attachment links to embeds (if any)
    if (attachmentLinks && attachmentLinks.length > 0) {
      const attachmentChunks = [];
      let currentChunk = "";

      // Create chunks of attachment links
      for (const [index, link] of attachmentLinks.entries()) {
        const line = `[Attachment ${index + 1}](${link})\n`;
        if (currentChunk.length + line.length > MAX_DESCRIPTION_LENGTH) {
          attachmentChunks.push(currentChunk);
          currentChunk = line;
        } else {
          currentChunk += line;
        }
      }
      if (currentChunk) {
        attachmentChunks.push(currentChunk);
      }

      // Add attachment chunks to embeds
      for (const chunk of attachmentChunks) {
        embeds.push({
          description: `**📎 Attachments:**\n${chunk}`,
        });
      }
    }

    // Add email content to embeds
    let content = emailTextContent;

    while (content.length > 0) {
      let chunk = content.substring(0, MAX_DESCRIPTION_LENGTH);
      const lastSpaceIndex = chunk.lastIndexOf(" ");
      if (lastSpaceIndex > -1 && lastSpaceIndex > chunk.length * 0.8) {
        chunk = chunk.substring(0, lastSpaceIndex);
      }

      content = content.substring(chunk.length).trim();

      // Check if the remaining content is small enough to append
      if (
        content.length > 0 &&
        content.length < 500 && // Threshold for small remaining content
        chunk.length + content.length <= MAX_DESCRIPTION_LENGTH
      ) {
        // Append remaining content to current chunk
        chunk += ' ' + content;
        content = '';
      }

      embeds.push({
        description: chunk,
      });
    }

    // Send embeds with batching
    await sendEmbedsWithBatching(channelId, embeds);

    // Send the text content as an attachment
    if (emailTextContent && emailTextContent.length > 0) {
      await sendTextAttachment(channelId, emailTextContent, attachmentLinks);
    }

    logToDiscord("INFO", `Sent email content to channel ID: ${channelId}`);
  } catch (error) {
    console.error("Error sending embed message:", error);
    logToDiscord("ERROR", `Error sending embed message: ${error.message}`);
    throw new Error("Failed to send an embed message to the member.");
  }
}

/**
 * Function to send log entries as embeds to the log channel with attachment
 */
async function sendLogEmbeds(detailedLog) {
  if (!LOG_CHANNEL_ID || logEntries.length === 0) return; // If no log channel is specified or no logs, skip

  try {
    const MAX_EMBEDS_PER_MESSAGE = 10;
    const MAX_EMBED_DESCRIPTION_LENGTH = 4096;
    const MAX_TOTAL_EMBED_SIZE = 6000; // Max total size of embeds per message

    const embedBatches = [];
    let currentBatch = [];
    let currentTotalEmbedSize = 0;

    for (const entry of logEntries) {
      const logMessage = `**[${entry.level}] ${entry.timestamp}**\n${entry.message}\n\n`;
      let description = logMessage;

      // Truncate if necessary
      if (description.length > MAX_EMBED_DESCRIPTION_LENGTH) {
        description = description.substring(0, MAX_EMBED_DESCRIPTION_LENGTH - 3) + '...';
      }

      const embed = {
        description,
        color: entry.level === "ERROR" ? 0xff0000 : entry.level === "WARN" ? 0xffa500 : 0x00ff00,
      };

      const embedSize = JSON.stringify(embed).length;

      // Check if adding this embed exceeds the total embed size limit or max embeds per message
      if (
        currentTotalEmbedSize + embedSize > MAX_TOTAL_EMBED_SIZE ||
        currentBatch.length >= MAX_EMBEDS_PER_MESSAGE
      ) {
        // Add current batch to batches
        embedBatches.push([...currentBatch]);
        // Reset batch and total size
        currentBatch = [];
        currentTotalEmbedSize = 0;
      }

      // Add embed to current batch
      currentBatch.push(embed);
      currentTotalEmbedSize += embedSize;
    }

    // Add any remaining embeds to batches
    if (currentBatch.length > 0) {
      embedBatches.push([...currentBatch]);
    }

    // Create the log text content
    const logTextContent = logEntries
      .map((entry) => `[${entry.level}] ${entry.timestamp} - ${entry.message}`)
      .join("\n");

    // Append detailed data to the log text content
    const detailedLogContent = `\n\nDetailed Data:\n${JSON.stringify(detailedLog, null, 2)}`;
    const fullLogTextContent = logTextContent + detailedLogContent;

    // Create a Blob for the log text attachment
    const logBlob = new Blob([fullLogTextContent], { type: "text/plain" });
    const logAttachment = {
      blob: logBlob,
      filename: `log_${new Date().toISOString()}.txt`,
    };

    // Send embed batches with attachment
    for (const [index, batch] of embedBatches.entries()) {
      const payload = { embeds: batch };

      // Attach the log file only once
      const attachments = index === 0 ? [logAttachment] : [];

      await sendEmbedBatch(LOG_CHANNEL_ID, payload, attachments);
    }
  } catch (error) {
    console.error("Error sending log embeds:", error);
  } finally {
    // Clear log entries after sending
    logEntries = [];
  }
}

/**
 * Main Handler
 */
export default {
  async email(event, env, ctx) {
    let parsedEmail;
    try {
      // Initialize environment variables with defaults or throw errors if required variables are missing
      if (!env.GUILD_ID) throw new Error("GUILD_ID environment variable is not defined.");
      if (!env.TOKEN) throw new Error("TOKEN environment variable is not defined.");
      if (!env.ATTACHMENTS_CHANNEL)
        throw new Error("ATTACHMENTS_CHANNEL environment variable is not defined.");

      GUILD_ID = env.GUILD_ID;
      TOKEN = env.TOKEN;
      ROLES_REQUIRED = env.ROLES_REQUIRED ? env.ROLES_REQUIRED.split(",") : [];

      if (env.CHANNEL_MAP) {
        CHANNEL_MAP = Object.fromEntries(
          env.CHANNEL_MAP.split(",").map((item) => item.trim().split(":"))
        );
      } else {
        CHANNEL_MAP = {};
      }

      ATTACHMENTS_CHANNEL = env.ATTACHMENTS_CHANNEL;
      LOG_CHANNEL_ID = env.LOG_CHANNEL_ID; // New environment variable for logging

      myHeaders.append("Authorization", `Bot ${TOKEN}`);
      logToDiscord("INFO", "Email event received.");

      // Log event details
      logToDiscord("INFO", `Event details:\nFrom: ${event.from}\nTo: ${event.to}`);

      parsedEmail = await parseEmail(event);
      const username = parsedEmail.to[0]?.address.split("@")[0] || "Unknown";
      logToDiscord("INFO", `Parsed email intended for username: ${username}`);

      // Log email details
      logToDiscord(
        "INFO",
        `Email details:\nSubject: ${parsedEmail.subject || "No Subject"}\nFrom: ${parsedEmail.from?.address || parsedEmail.from?.[0]?.address || event.from || "Unknown"
        }\nTo: ${parsedEmail.to?.map((addr) => addr.address).join(", ") || event.to || "Unknown"
        }\nDate: ${parsedEmail.date || new Date().toISOString()
        }\nAttachments: ${parsedEmail.attachments?.length || 0}`
      );

      // Sanitize and log parsedEmail data
      const sanitizedParsedEmail = sanitizeData(parsedEmail);
      logToDiscord("INFO", `Parsed Email Data:\n${JSON.stringify(sanitizedParsedEmail, null, 2)}`);

      // Log ctx data
      logToDiscord("INFO", `Context Data:\n${JSON.stringify(ctx, null, 2)}`);

      // Handle attachments
      let attachmentLinks = [];
      if (parsedEmail.attachments && parsedEmail.attachments.length > 0) {
        attachmentLinks = await handleAttachments(parsedEmail.attachments);
      }

      // Determine target channel
      let targetChannelId;
      if (CHANNEL_MAP[username]) {
        targetChannelId = CHANNEL_MAP[username];
        logToDiscord("INFO", `Using channel from CHANNEL_MAP: ${targetChannelId}`);
      } else {
        const member = await findDiscordMember(username);

        if (member) {
          if (!hasRequiredRoles(member)) {
            const errorMsg = "Member does not have the required role(s).";
            logToDiscord("WARN", errorMsg);
            throw new Error(errorMsg);
          }
          const dm = await createDmChannel(member.user.id);
          targetChannelId = dm.id;
          logToDiscord("INFO", `DM channel created: ${targetChannelId}`);
        } else if (CHANNEL_MAP["others"]) {
          targetChannelId = CHANNEL_MAP["others"];
          logToDiscord("INFO", `Using 'others' channel from CHANNEL_MAP: ${targetChannelId}`);
        } else {
          const errorMsg = "No target channel found for the recipient.";
          logToDiscord("ERROR", errorMsg);
          throw new Error(errorMsg);
        }
      }

      // Send message to Discord
      await sendEmbedMessage(targetChannelId, parsedEmail, event, attachmentLinks);

      // Send auto-reply
      await sendAutoReply(event, parsedEmail);

      logToDiscord("INFO", "Email processing completed successfully.");
    } catch (error) {
      console.error("Error handling email event:", error);
      logToDiscord("ERROR", `Error handling email event: ${error.message}`);
      if (parsedEmail) {
        await sendAutoReply(event, parsedEmail, error.message);
      }
    } finally {
      // Prepare detailed log data
      const detailedLog = {
        parsedEmail: sanitizeData(parsedEmail) || {},
        event: {
          from: event.from,
          to: event.to,
          headers: event.headers,
        },
        ctx: ctx || {},
      };

      // Send all accumulated logs as embeds with attachment
      await sendLogEmbeds(detailedLog);
    }
  },
};
