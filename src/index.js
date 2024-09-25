// Import necessary modules
import { Buffer } from "node:buffer";
globalThis.Buffer = Buffer;

import { EmailMessage } from "cloudflare:email";
import { createMimeMessage } from "mimetext";
import PostalMime from "postal-mime";

// Discord API version and base URL
const API_VERSION = 10;
const DISCORD_API_URL = `https://discord.com/api/v${API_VERSION}`;

// Variables to hold environment variables
var GUILD_ID, TOKEN, ROLES_REQUIRED, CHANNEL_MAP, ATTACHMENTS_CHANNEL;

// Headers and request options
let myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");

const requestOptions = {
  headers: myHeaders,
  redirect: "follow",
};

// Utility function to convert a stream to an ArrayBuffer
async function streamToArrayBuffer(stream, streamSize) {
  try {
    const result = new Uint8Array(streamSize);
    let bytesRead = 0;
    const reader = stream.getReader();

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      result.set(value, bytesRead);
      bytesRead += value.length;
    }

    return result;
  } catch (error) {
    console.error("Error converting stream to ArrayBuffer:", error);
    throw new Error("Failed to process email stream.");
  }
}

// Function to parse the email using PostalMime
async function parseEmail(event) {
  try {
    const rawEmail = await streamToArrayBuffer(event.raw, event.rawSize);
    const parser = new PostalMime();
    return await parser.parse(rawEmail);
  } catch (error) {
    console.error("Error parsing email:", error);
    throw new Error("Failed to parse email content.");
  }
}

// Function to find Discord member by username
async function findDiscordMember(username) {
  try {
    const fetchMemberURL = `${DISCORD_API_URL}/guilds/${GUILD_ID}/members/search?query=${encodeURIComponent(
      username
    )}&limit=1000`;
    const response = await fetchWithRateLimit(fetchMemberURL, requestOptions);
    if (!response.ok) {
      throw new Error(
        `Failed to fetch Discord members. Status: ${response.status}`
      );
    }
    const members = await response.json();
    return members.find((mem) => mem.user.username === username);
  } catch (error) {
    console.error("Error fetching Discord member:", error);
    throw new Error(`Failed to find Discord member with username: ${username}`);
  }
}

// Function to check if the member has the required role(s)
async function hasRequiredRoles(member) {
  if (!member) return false;
  const memberRoles = member.roles; // This is an array of role IDs the member has

  if (ROLES_REQUIRED.length === 0) return true; // No roles required, so return true

  // Check if the member has at least one of the required roles
  return ROLES_REQUIRED.some((role) => memberRoles.includes(role));
}

// Function to create a DM channel with the user
async function createDmChannel(recipientId) {
  try {
    const createDmURL = `${DISCORD_API_URL}/users/@me/channels`;
    const response = await fetchWithRateLimit(createDmURL, {
      method: "POST",
      body: JSON.stringify({ recipient_id: recipientId }),
      ...requestOptions,
    });
    if (!response.ok) {
      throw new Error(
        `Failed to create DM channel. Status: ${response.status}`
      );
    }
    return await response.json();
  } catch (error) {
    console.error("Error creating DM channel:", error);
    throw new Error(
      "Failed to create a direct message channel with the member."
    );
  }
}

// Function to send an auto-reply email, optionally with an error message
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

    const msg = createMimeMessage();
    msg.setSender({ name: "Auto-replier", addr: event.to });
    msg.setRecipient(event.from);
    msg.setSubject(subject);

    if (parsedEmail?.messageId) {
      msg.setHeader("In-Reply-To", parsedEmail.messageId);
    }

    msg.addMessage({
      contentType: "text/html",
      data: messageData,
    });

    const message = new EmailMessage(event.to, event.from, msg.asRaw());
    await event.reply(message);
  } catch (error) {
    console.error("Error sending auto-reply:", error);
  }
}

// Function to fetch guild data from Discord API
async function fetchGuildData(guildId) {
  try {
    const url = `${DISCORD_API_URL}/guilds/${guildId}`;
    const response = await fetchWithRateLimit(url, {
      method: "GET",
      headers: myHeaders,
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch guild data. Status: ${response.status}`);
    }

    const guildData = await response.json();
    return guildData;
  } catch (error) {
    console.error("Error fetching guild data:", error);
    throw new Error("Failed to fetch guild information.");
  }
}

// Function to get the guild icon URL from fetched guild data
function getGuildIconURL(guildData) {
  if (guildData.icon) {
    const fileExtension = guildData.icon.startsWith("a_") ? ".gif" : ".png";
    return `https://cdn.discordapp.com/icons/${guildData.id}/${guildData.icon}${fileExtension}`;
  }
  return null; // If no icon is set
}

// Function to safely truncate a string to a specific length with an ellipsis if needed
function truncateText(text, maxLength) {
  return text.length > maxLength
    ? text.substring(0, maxLength - 3) + "..."
    : text;
}

// Improved function to extract text from HTML content
function extractTextFromHtml(htmlContent) {
  try {
    // Remove script and style tags and their content
    htmlContent = htmlContent.replace(
      /<script[^>]*>([\s\S]*?)<\/script>/gi,
      ""
    );
    htmlContent = htmlContent.replace(
      /<style[^>]*>([\s\S]*?)<\/style>/gi,
      ""
    );
    // Replace line breaks and paragraph tags with newlines
    htmlContent = htmlContent.replace(/<(br|\/p|p)[^>]*>/gi, "\n");
    // Remove all remaining HTML tags
    htmlContent = htmlContent.replace(/<[^>]+>/g, "");
    // Replace multiple spaces and newlines with a single space
    htmlContent = htmlContent.replace(/\s+/g, " ").trim();
    // Decode HTML entities
    const decodedText = decodeHtmlEntities(htmlContent);
    return decodedText;
  } catch (error) {
    console.error("Error extracting text from HTML:", error);
    return "(Error extracting text content)";
  }
}

// Function to decode HTML entities
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
  return text.replace(/&([^;]+);/g, (match, entity) => {
    if (entities[entity]) {
      return entities[entity];
    }
    return match;
  });
}

// Function to extract links and images from HTML content
function extractLinksAndImages(htmlContent) {
  const httpLinks = [];
  const mailtoLinks = [];
  const telLinks = [];
  const images = [];

  // Extract links
  const linkRegex = /<a[^>]*href="([^"]*)"[^>]*>(.*?)<\/a>/gi;
  let match;
  while ((match = linkRegex.exec(htmlContent)) !== null) {
    const href = match[1];
    const text = match[2];
    if (href.startsWith("http://") || href.startsWith("https://")) {
      httpLinks.push({ href, text });
    } else if (href.startsWith("mailto:")) {
      mailtoLinks.push({ href, text });
    } else if (href.startsWith("tel:")) {
      telLinks.push({ href, text });
    } else {
      // Handle other schemes if necessary
      console.warn(`Unsupported link scheme: ${href}`);
    }
  }

  // Extract images
  const imgRegex = /<img[^>]*src="([^"]*)"[^>]*>/gi;
  while ((match = imgRegex.exec(htmlContent)) !== null) {
    images.push(match[1]);
  }

  return { httpLinks, mailtoLinks, telLinks, images };
}

// Function to send multiple embeds in Discord messages
async function sendEmbedsWithBatching(channelId, embeds) {
  const maxEmbedsPerMessage = 10;
  const maxTotalEmbedSize = 6000; // Maximum total size of embeds per message
  const messages = [];

  let currentBatch = [];
  let currentTotalSize = 0;

  for (const embed of embeds) {
    const embedSize = JSON.stringify(embed).length;

    // If adding this embed would exceed the size limit or embed count limit, send the current batch
    if (
      currentTotalSize + embedSize > maxTotalEmbedSize ||
      currentBatch.length >= maxEmbedsPerMessage
    ) {
      // Send current batch
      const payload = { embeds: currentBatch };
      await sendEmbedBatch(channelId, payload);
      messages.push(payload);
      // Reset batch
      currentBatch = [];
      currentTotalSize = 0;
    }

    // Add embed to current batch
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

// Helper function to send a batch of embeds
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
    throw new Error(
      `Failed to send message. Status: ${response.status}. Error: ${errorText}`
    );
  }

  return await response.json();
}

// Function to send the text attachment in a separate message
async function sendTextAttachment(channelId, textContent, attachmentLinks) {
  try {
    const formData = new FormData();

    // Combine text content and attachments data
    let fullTextContent = textContent;

    if (attachmentLinks && attachmentLinks.length > 0) {
      const attachmentsText = attachmentLinks
        .map((link, index) => `Attachment ${index + 1}: ${link}`)
        .join("\n");
      fullTextContent += `\n\nAttachments:\n${attachmentsText}`;
    }

    // Create a Blob from the full text content
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

// Function to send an embed message in a Discord channel
async function sendEmbedMessage(channelId, parsedEmail, event, attachmentLinks) {
  try {
    // Fetch guild data to get the icon dynamically
    const guildData = await fetchGuildData(GUILD_ID);
    const thumbnailUrl = getGuildIconURL(guildData);

    // Prepare "From" and "To" information
    const fromField = event.from;
    const toField = parsedEmail.to.map((addr) => addr.address).join(", ");

    // Extract text content from parsedEmail
    let emailTextContent = parsedEmail.text;

    if (!emailTextContent && parsedEmail.html) {
      emailTextContent = extractTextFromHtml(parsedEmail.html);
    }

    // Provide a fallback if no text content is available
    emailTextContent = emailTextContent || "(No text content)";

    // Maximum allowed length for description field
    const maxDescriptionLength = 4096;

    // Construct the initial embed description
    let descriptionParts = [];

    // Add "From" and "To" information
    descriptionParts.push(`**👤 From:** ${fromField}`);
    descriptionParts.push(`**📩 To:** ${toField}`);

    // Combine the initial description
    let initialDescription = descriptionParts.join("\n\n");

    // Ensure initialDescription does not exceed maxDescriptionLength
    if (initialDescription.length > maxDescriptionLength) {
      initialDescription = initialDescription.substring(0, maxDescriptionLength);
    }

    // Now, create the first embed
    const title = `📧 ${parsedEmail.subject || "New Email Received"}`;
    const footerText = "📬 Sent via Clanflare Email System";
    const timestamp = new Date().toISOString();

    // Prepare the first embed
    let firstEmbed = {
      title: title,
      description: initialDescription, // We will add content after adjusting
      footer: {
        text: footerText,
      },
      timestamp: timestamp,
    };

    let embeds = [firstEmbed];

    // Prepare attachment links to be added to embeds
    if (attachmentLinks && attachmentLinks.length > 0) {
      const attachmentTextChunks = [];
      let currentChunk = "";

      // Create chunks of attachment links that fit within the description limit
      for (const [index, link] of attachmentLinks.entries()) {
        const attachmentLine = `[Attachment ${index + 1}](${link})\n`;
        if ((currentChunk + attachmentLine).length > maxDescriptionLength) {
          attachmentTextChunks.push(currentChunk);
          currentChunk = attachmentLine;
        } else {
          currentChunk += attachmentLine;
        }
      }
      if (currentChunk) {
        attachmentTextChunks.push(currentChunk);
      }

      // Add attachment chunks to embeds
      for (const chunk of attachmentTextChunks) {
        // If there's room in the last embed's description, add it there
        let lastEmbed = embeds[embeds.length - 1];
        if (
          lastEmbed.description &&
          lastEmbed.description.length +
            `\n\n**📎 Attachments:**\n${chunk}`.length <=
            maxDescriptionLength
        ) {
          lastEmbed.description += `\n\n**📎 Attachments:**\n${chunk}`;
        } else {
          // Create a new embed for attachments
          embeds.push({
            description: `**📎 Attachments:**\n${chunk}`,
          });
        }
      }
    }

    // Add email content to embeds
    let content = emailTextContent;

    while (content.length > 0) {
      let chunk = content.substring(0, maxDescriptionLength);
      // Ensure we don't split in the middle of a word
      const lastSpaceIndex = chunk.lastIndexOf(" ");
      if (lastSpaceIndex > -1 && lastSpaceIndex > chunk.length * 0.8) {
        chunk = chunk.substring(0, lastSpaceIndex);
      }

      embeds.push({
        description: chunk,
      });

      content = content.substring(chunk.length).trim();
    }

    // Log the number of embeds
    console.log(`Total number of embeds: ${embeds.length}`);

    // Send embeds with batching
    await sendEmbedsWithBatching(channelId, embeds);

    // Send the text attachment in a separate message
    if (emailTextContent && emailTextContent.length > 0) {
      await sendTextAttachment(channelId, emailTextContent, attachmentLinks);
    }

    return; // No need to return a response here
  } catch (error) {
    console.error("Error sending embed message:", error);
    throw new Error("Failed to send an embed message to the member.");
  }
}

// Function to handle rate limits in fetch requests
async function fetchWithRateLimit(url, options, retries = 3) {
  let response = await fetch(url, options);

  while (response.status === 429 && retries > 0) {
    const retryAfter = response.headers.get("Retry-After");
    const retryAfterMs = retryAfter ? parseFloat(retryAfter) * 1000 : 1000; // Default to 1 second
    console.warn(`Rate limited. Retrying after ${retryAfterMs}ms`);
    await new Promise((resolve) => setTimeout(resolve, retryAfterMs));
    // Retry the request
    response = await fetch(url, options);
    retries--;
  }

  if (response.status === 429) {
    throw new Error("Rate limit exceeded, maximum retries reached.");
  }

  return response;
}

// Function to handle attachments
async function handleAttachments(attachments) {
  const attachmentLinks = [];

  // Batch attachments into groups of up to 10 (Discord's limit per message)
  const batches = [];
  const batchSize = 10;

  for (let i = 0; i < attachments.length; i += batchSize) {
    batches.push(attachments.slice(i, i + batchSize));
  }

  for (const batch of batches) {
    // Upload the batch and get the attachment URLs
    const links = await uploadAttachmentsToChannel(ATTACHMENTS_CHANNEL, batch);
    if (links && links.length > 0) {
      attachmentLinks.push(...links);
    }
  }

  return attachmentLinks;
}

// Function to upload multiple attachments to a specific channel and return their URLs
async function uploadAttachmentsToChannel(channelId, attachments) {
  try {
    const formData = new FormData();

    // Append up to 10 attachments
    attachments.forEach((attachment, index) => {
      const blob = new Blob([attachment.content], {
        type: attachment.contentType,
      });
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

    // Get the attachment URLs from the message data
    if (messageData.attachments && messageData.attachments.length > 0) {
      const attachmentUrls = messageData.attachments.map(
        (attachment) => attachment.url
      );
      return attachmentUrls;
    } else {
      console.error("No attachment URLs found in message data.");
      return [];
    }
  } catch (error) {
    console.error("Error uploading attachments:", error);
    return [];
  }
}

// Main function to handle the email event
export default {
  async email(event, env, ctx) {
    GUILD_ID = env.GUILD_ID;
    TOKEN = env.TOKEN;
    ROLES_REQUIRED =
      env.ROLES_REQUIRED === "" ? [] : env.ROLES_REQUIRED.split(",");
    CHANNEL_MAP = Object.fromEntries(
      env.CHANNEL_MAP.split(",").map((item) => item.trim().split(":"))
    );
    ATTACHMENTS_CHANNEL = env.ATTACHMENTS_CHANNEL;

    myHeaders.append("Authorization", `Bot ${TOKEN}`);
    let parsedEmail = null;
    try {
      parsedEmail = await parseEmail(event);
      const username = parsedEmail.to[0].address.split("@")[0];

      // Handle attachments first and get the attachment links
      let attachmentLinks = [];
      if (parsedEmail.attachments && parsedEmail.attachments.length > 0) {
        attachmentLinks = await handleAttachments(parsedEmail.attachments);
      }

      let targetChannelId;
      if (CHANNEL_MAP[username]) {
        targetChannelId = CHANNEL_MAP[username];
      } else {
        const member = await findDiscordMember(username);

        if (member) {
          if (!(await hasRequiredRoles(member))) {
            throw new Error(`Member does not have the required role(s).`);
          }
          const dm = await createDmChannel(member.user.id);
          targetChannelId = dm.id;
        } else {
          targetChannelId = CHANNEL_MAP["others"];
        }
      }

      await sendEmbedMessage(
        targetChannelId,
        parsedEmail,
        event,
        attachmentLinks
      );

      await sendAutoReply(event, parsedEmail);
    } catch (error) {
      console.error("Error handling email event:", error);
      await sendAutoReply(event, parsedEmail, error.message);
    }
  },
};
