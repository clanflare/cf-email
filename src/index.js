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
let GUILD_ID, TOKEN, ROLES_REQUIRED, CHANNEL_MAP, ATTACHMENTS_CHANNEL;

// Headers and request options
const myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");

// Rate limiting configuration
const MAX_RETRIES = 3;
const DEFAULT_RETRY_AFTER_MS = 1000; // 1 second

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
    throw new Error("Rate limit exceeded, maximum retries reached.");
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
 * Email Handling Functions
 */

// Parse the email using PostalMime
async function parseEmail(event) {
  try {
    const rawEmail = await streamToArrayBuffer(event.raw);
    const parser = new PostalMime();
    return await parser.parse(rawEmail);
  } catch (error) {
    console.error("Error parsing email:", error);
    throw new Error("Failed to parse email content.");
  }
}

// Send an auto-reply email
async function sendAutoReply(event, parsedEmail, errorMsg = null) {
  try {
    const timestamp = new Date().toISOString();
    const subject = `Re: ${parsedEmail.subject || "No Subject"}`;
    const originalRecipient = parsedEmail.to[0].address.split("@")[0];
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

    const fromAddress = parsedEmail.from.address;
    const toAddress = parsedEmail.to[0].address;

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
    } catch (replyError) {
      console.warn("event.reply() failed, attempting to send a new email:", replyError);

      // If event.reply() fails, send a new email instead
      await event.send(message);
    }
  } catch (error) {
    console.error("Error sending auto-reply:", error);
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

    return await response.json();
  } catch (error) {
    console.error("Error fetching guild data:", error);
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
    return members.find((mem) => mem.user.username === username);
  } catch (error) {
    console.error("Error fetching Discord member:", error);
    throw new Error(`Failed to find Discord member with username: ${username}`);
  }
}

// Check if the member has the required role(s)
function hasRequiredRoles(member) {
  if (!member) return false;
  const memberRoles = member.roles; // Array of role IDs the member has

  if (ROLES_REQUIRED.length === 0) return true; // No roles required

  return ROLES_REQUIRED.some((role) => memberRoles.includes(role));
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

    return await response.json();
  } catch (error) {
    console.error("Error creating DM channel:", error);
    throw new Error("Failed to create a direct message channel with the member.");
  }
}

// Send a batch of embeds to a Discord channel
async function sendEmbedBatch(channelId, payload) {
  const response = await fetchWithRateLimit(
    `${DISCORD_API_URL}/channels/${channelId}/messages`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bot ${TOKEN}`,
      },
      body: JSON.stringify(payload),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to send message. Status: ${response.status}. Error: ${errorText}`);
  }

  return await response.json();
}

// Send multiple embeds with batching
async function sendEmbedsWithBatching(channelId, embeds) {
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
    await sendEmbedBatch(channelId, payload);
    messages.push(payload);
  }

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
      throw new Error(
        `Failed to send message with attachment. Status: ${response.status}. Error: ${errorText}`
      );
    }

    return await response.json();
  } catch (error) {
    console.error("Error sending text attachment:", error);
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
      throw new Error(
        `Failed to upload attachments. Status: ${response.status}. Error: ${errorText}`
      );
    }

    const messageData = await response.json();

    // Extract attachment URLs
    if (messageData.attachments && messageData.attachments.length > 0) {
      return messageData.attachments.map((attachment) => attachment.url);
    } else {
      console.error("No attachment URLs found in message data.");
      return [];
    }
  } catch (error) {
    console.error("Error uploading attachments:", error);
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

  return attachmentLinks;
}

// Send the email content as embeds to Discord
async function sendEmbedMessage(channelId, parsedEmail, event, attachmentLinks) {
  try {
    // Fetch guild data and icon
    const guildData = await fetchGuildData();
    const thumbnailUrl = getGuildIconURL(guildData);

    // Prepare email fields
    const fromField = event.from;
    const toField = parsedEmail.to.map((addr) => addr.address).join(", ");

    // Extract email text content
    let emailTextContent = parsedEmail.text || extractTextFromHtml(parsedEmail.html);
    emailTextContent = emailTextContent || "(No text content)";

    // Maximum lengths
    const MAX_DESCRIPTION_LENGTH = 4096;

    // Construct the initial embed
    const title = `📧 ${parsedEmail.subject || "New Email Received"}`;
    const footerText = "📬 Sent via Clanflare Email System";
    const timestamp = new Date().toISOString();

    let embeds = [
      {
        title,
        description: truncateText(`**👤 From:** ${fromField}\n\n**📩 To:** ${toField}`, MAX_DESCRIPTION_LENGTH),
        footer: { text: footerText },
        timestamp,
        thumbnail: { url: thumbnailUrl },
      },
    ];

    // Add attachment links to embeds
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

      embeds.push({
        description: chunk,
      });

      content = content.substring(chunk.length).trim();
    }

    // Send embeds with batching
    await sendEmbedsWithBatching(channelId, embeds);

    // Send the text content as an attachment
    if (emailTextContent && emailTextContent.length > 0) {
      await sendTextAttachment(channelId, emailTextContent, attachmentLinks);
    }
  } catch (error) {
    console.error("Error sending embed message:", error);
    throw new Error("Failed to send an embed message to the member.");
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
      if (!env.ATTACHMENTS_CHANNEL) throw new Error("ATTACHMENTS_CHANNEL environment variable is not defined.");

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

      myHeaders.append("Authorization", `Bot ${TOKEN}`);
      parsedEmail = await parseEmail(event);
      const username = parsedEmail.to[0].address.split("@")[0];

      // Handle attachments
      let attachmentLinks = [];
      if (parsedEmail.attachments && parsedEmail.attachments.length > 0) {
        attachmentLinks = await handleAttachments(parsedEmail.attachments);
      }

      // Determine target channel
      let targetChannelId;
      if (CHANNEL_MAP[username]) {
        targetChannelId = CHANNEL_MAP[username];
      } else {
        const member = await findDiscordMember(username);

        if (member) {
          if (!hasRequiredRoles(member)) {
            throw new Error("Member does not have the required role(s).");
          }
          const dm = await createDmChannel(member.user.id);
          targetChannelId = dm.id;
        } else if (CHANNEL_MAP["others"]) {
          targetChannelId = CHANNEL_MAP["others"];
        } else {
          throw new Error("No target channel found for the recipient.");
        }
      }

      // Send message to Discord
      await sendEmbedMessage(targetChannelId, parsedEmail, event, attachmentLinks);

      // Send auto-reply
      await sendAutoReply(event, parsedEmail);
    } catch (error) {
      console.error("Error handling email event:", error);
      await sendAutoReply(event, parsedEmail, error.message);
    }
  },
};
