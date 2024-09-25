// Import necessary modules
import { Buffer } from "node:buffer";
globalThis.Buffer = Buffer;

import { EmailMessage } from "cloudflare:email";
import { createMimeMessage } from "mimetext";
const PostalMime = require("postal-mime");

// Discord API version and base URL
const API_VERSION = 10;
const DISCORD_API_URL = `https://discord.com/api/v${API_VERSION}`;

// Variables to hold environment variables
var GUILD_ID, TOKEN, ROLES_REQUIRED, CHANNEL_MAP;

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
    const parser = new PostalMime.default();
    return await parser.parse(rawEmail);
  } catch (error) {
    console.error("Error parsing email:", error);
    throw new Error("Failed to parse email content.");
  }
}

// Function to find Discord member by username
async function findDiscordMember(username) {
  try {
    const fetchMemberURL = `${DISCORD_API_URL}/guilds/${GUILD_ID}/members/search?query=${encodeURIComponent(username)}&limit=1000`;
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
    const subject = `Re: ${parsedEmail.subject}`;
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
        <p>Thank you for your email. This is an automated reply to confirm that your email with the subject "<strong>${parsedEmail.subject}</strong>" has been successfully delivered to the recipient: <strong>${originalRecipient}</strong>.</p>
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

// Function to extract text from HTML content
function extractTextFromHtml(htmlContent) {
  // Remove script and style tags and their content
  htmlContent = htmlContent.replace(/<script[^>]*>([\s\S]*?)<\/script>/gi, '');
  htmlContent = htmlContent.replace(/<style[^>]*>([\s\S]*?)<\/style>/gi, '');
  // Replace line breaks and paragraph tags with newlines
  htmlContent = htmlContent.replace(/<(br|\/p|p)[^>]*>/gi, '\n');
  // Remove all remaining HTML tags
  const textContent = htmlContent.replace(/<[^>]+>/g, '');
  // Decode HTML entities
  const decodedText = decodeHtmlEntities(textContent);
  return decodedText;
}

// Function to decode HTML entities
function decodeHtmlEntities(text) {
  const entities = {
    nbsp: ' ',
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'",
    '#39': "'",
  };
  return text.replace(/&([^;]+);/g, (match, entity) => {
    if (entities[entity]) {
      return entities[entity];
    }
    return match;
  });
}

// Function to send an embed message in a Discord channel
async function sendEmbedMessage(channelId, parsedEmail, event) {
  try {
    // Fetch guild data to get the icon dynamically
    const guildData = await fetchGuildData(GUILD_ID);
    const thumbnailUrl = getGuildIconURL(guildData);

    const sendMessageURL = `${DISCORD_API_URL}/channels/${channelId}/messages`;

    // Extract text content from parsedEmail
    let emailTextContent = parsedEmail.text;

    if (!emailTextContent && parsedEmail.html) {
      emailTextContent = extractTextFromHtml(parsedEmail.html);
    }

    // Provide a fallback if no text content is available
    emailTextContent = emailTextContent || "(No text content)";

    // Truncate content to respect Discord embed character limits
    const title = truncateText(`📧 New Email Received`, 256);
    const description = truncateText(emailTextContent, 4096);
    const fromField = truncateText(event.from, 256);
    const toField = truncateText(
      parsedEmail.to.map((addr) => addr.address).join(", "),
      1024
    );
    const subjectField = truncateText(
      parsedEmail.subject || "(No subject)",
      1024
    );
    const footerText = truncateText("📬 Sent via Clanflare Email System", 2048);

    const embed = {
      title: title,
      description: description,
      color: 0x5865f2, // Discord's blurple color
      fields: [
        {
          name: "👤 **From**",
          value: fromField,
          inline: true,
        },
        {
          name: "📩 **To**",
          value: toField,
          inline: true,
        },
        {
          name: "📝 **Subject**",
          value: subjectField,
          inline: false,
        },
      ],
      footer: {
        text: footerText,
      },
      timestamp: new Date().toISOString(),
      thumbnail: {
        url: thumbnailUrl || "https://cdn.discordapp.com/embed/avatars/0.png", // Default avatar if no icon is set
      },
    };

    const response = await fetchWithRateLimit(sendMessageURL, {
      method: "POST",
      body: JSON.stringify({ embeds: [embed] }),
      ...requestOptions,
    });

    if (!response.ok) {
      throw new Error(
        `Failed to send embed message. Status: ${response.status}`
      );
    }

    // Check if description was truncated, and send full content as a follow-up if necessary
    if (description !== emailTextContent) {
      await sendFullTextMessage(channelId, emailTextContent);
    }

    // Handle attachments if any
    if (parsedEmail.attachments && parsedEmail.attachments.length > 0) {
      await handleAttachments(channelId, parsedEmail.attachments);
    }

    return await response.json();
  } catch (error) {
    console.error("Error sending embed message:", error);
    throw new Error("Failed to send an embed message to the member.");
  }
}

// Function to send the full text message in chunks
async function sendFullTextMessage(channelId, textContent) {
  try {
    const sendMessageURL = `${DISCORD_API_URL}/channels/${channelId}/messages`;
    const maxMessageLength = 2000; // Discord's message character limit
    let index = 0;
    const totalLength = textContent.length;

    while (index < totalLength) {
      const chunk = textContent.substring(index, index + maxMessageLength - 6); // Adjust for code block markers
      const formattedChunk = `\`\`\`\n${chunk}\n\`\`\``; // Wrap in code block

      const response = await fetchWithRateLimit(sendMessageURL, {
        method: "POST",
        body: JSON.stringify({ content: formattedChunk }),
        ...requestOptions,
      });
      if (!response.ok) {
        throw new Error(
          `Failed to send message chunk. Status: ${response.status}`
        );
      }
      index += maxMessageLength - 6;
    }
  } catch (error) {
    console.error("Error sending full text message:", error);
    throw new Error("Failed to send the full message content.");
  }
}

// Function to handle rate limits in fetch requests
async function fetchWithRateLimit(url, options) {
  let response = await fetch(url, options);

  if (response.status === 429) {
    const retryAfter = response.headers.get('Retry-After');
    const retryAfterMs = retryAfter ? parseFloat(retryAfter) * 1000 : 1000; // Default to 1 second
    console.warn(`Rate limited. Retrying after ${retryAfterMs}ms`);
    await new Promise((resolve) => setTimeout(resolve, retryAfterMs));
    // Retry the request
    response = await fetch(url, options);
  }

  return response;
}

// Function to handle attachments
async function handleAttachments(channelId, attachments) {
  for (const attachment of attachments) {
    // Check if the attachment is an image or other type
    if (attachment.contentType.startsWith('image/') || attachment.contentType.startsWith('application/')) {
      // Send the attachment as a file message in Discord
      await sendFileMessage(channelId, attachment);
    } else {
      // Handle other types of attachments as needed
      // For example, send a message indicating an unsupported attachment was received
      console.warn(`Unsupported attachment type: ${attachment.contentType}`);
    }
  }
}

// Function to send file messages to Discord
async function sendFileMessage(channelId, attachment) {
  const formData = new FormData();
  formData.append('files[0]', new Blob([attachment.content], { type: attachment.contentType }), attachment.filename);

  const sendMessageURL = `${DISCORD_API_URL}/channels/${channelId}/messages`;

  const response = await fetchWithRateLimit(sendMessageURL, {
    method: 'POST',
    headers: {
      'Authorization': `Bot ${TOKEN}`,
    },
    body: formData,
  });

  if (!response.ok) {
    throw new Error(`Failed to send file message. Status: ${response.status}`);
  }
}

// Main function to handle the email event
export default {
  async email(event, env, ctx) {
    GUILD_ID = env.GUILD_ID;
    TOKEN = env.TOKEN;
    ROLES_REQUIRED = env.ROLES_REQUIRED === "" ? [] : env.ROLES_REQUIRED.split(",");
    CHANNEL_MAP = Object.fromEntries(
      env.CHANNEL_MAP.split(",").map((item) => item.trim().split(":"))
    );

    myHeaders.append("Authorization", `Bot ${TOKEN}`);
    let parsedEmail = null;
    try {
      parsedEmail = await parseEmail(event);
      const username = parsedEmail.to[0].address.split("@")[0];
      if (CHANNEL_MAP[username]) {
        await sendEmbedMessage(CHANNEL_MAP[username], parsedEmail, event);
      } else {
        const member = await findDiscordMember(username);

        if (member) {
          if (!(await hasRequiredRoles(member))) {
            throw new Error(`Member does not have the required role(s).`);
          }
          const dm = await createDmChannel(member.user.id);
          await sendEmbedMessage(dm.id, parsedEmail, event);
        } else {
          await sendEmbedMessage(CHANNEL_MAP["others"], parsedEmail, event);
        }
      }
      await sendAutoReply(event, parsedEmail);
    } catch (error) {
      console.error("Error handling email event:", error);
      await sendAutoReply(event, parsedEmail, error.message);
    }
  },
};
