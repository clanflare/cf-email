// Import necessary modules
import { Buffer } from "node:buffer";
globalThis.Buffer = Buffer;

import { EmailMessage } from "cloudflare:email";
import { createMimeMessage } from "mimetext";
import PostalMime from "postal-mime";

/**
 * Configuration Constants
 */
const CONFIG = {
  API_VERSION: 10,
  DISCORD_API_URL: '', // Will be set after API_VERSION is known
  MAX_RETRIES: 3,
  DEFAULT_RETRY_AFTER_MS: 1000, // 1 second
  MAX_EMBEDS_PER_MESSAGE: 10,
  MAX_TOTAL_EMBED_SIZE: 6000, // Max total size of embeds per message
  MAX_DESCRIPTION_LENGTH: 4096,
  ATTACHMENT_BATCH_SIZE: 10,
  SMALL_CONTENT_THRESHOLD: 500, // Threshold for small remaining content in email
  MIN_CHUNK_MERGE_THRESHOLD: 100, // Threshold to merge small last chunk into previous embed
};
CONFIG.DISCORD_API_URL = `https://discord.com/api/v${CONFIG.API_VERSION}`;

// Environment variables (to be initialized in the main function)
let GUILD_ID, TOKEN, ROLES_REQUIRED, CHANNEL_MAP, ATTACHMENTS_CHANNEL, LOG_CHANNEL_ID;

// Headers and request options
const myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");

/**
 * Utility Functions
 */
const Utils = {
  // Convert a stream to an ArrayBuffer
  async streamToArrayBuffer(stream) {
    try {
      const arrayBuffer = await new Response(stream).arrayBuffer();
      return new Uint8Array(arrayBuffer);
    } catch (error) {
      console.error("Error converting stream to ArrayBuffer:", error);
      await Logger.logToDiscord("ERROR", `Error converting stream to ArrayBuffer: ${error.message}`);
      throw new Error("Failed to process email stream.");
    }
  },

  // Fetch with rate limit handling
  async fetchWithRateLimit(url, options, retries = CONFIG.MAX_RETRIES) {
    let response = await fetch(url, options);

    while (response.status === 429 && retries > 0) {
      const retryAfter = response.headers.get("Retry-After");
      const retryAfterMs = retryAfter
        ? parseFloat(retryAfter) * 1000
        : CONFIG.DEFAULT_RETRY_AFTER_MS;
      console.warn(`Rate limited. Retrying after ${retryAfterMs}ms`);
      await new Promise((resolve) => setTimeout(resolve, retryAfterMs));

      response = await fetch(url, options);
      retries--;
    }

    if (response.status === 429) {
      const errorMsg = "Rate limit exceeded, maximum retries reached.";
      console.error(errorMsg);
      await Logger.logToDiscord("ERROR", errorMsg);
      throw new Error(errorMsg);
    }

    return response;
  },

  // Truncate text safely
  truncateText(text, maxLength) {
    return text.length > maxLength ? `${text.substring(0, maxLength - 3)}...` : text;
  },

  // Decode HTML entities
  decodeHtmlEntities(text) {
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
  },

  // Extract text from HTML content
  extractTextFromHtml(htmlContent) {
    try {
      let text = htmlContent
        .replace(/<script[^>]*>([\s\S]*?)<\/script>/gi, "")
        .replace(/<style[^>]*>([\s\S]*?)<\/style>/gi, "")
        .replace(/<(br|\/p|p)[^>]*>/gi, "\n")
        .replace(/<[^>]+>/g, "")
        .replace(/\s+/g, " ")
        .trim();
      return Utils.decodeHtmlEntities(text);
    } catch (error) {
      console.error("Error extracting text from HTML:", error);
      return "(Error extracting text content)";
    }
  },

  // Sanitize data to avoid logging sensitive information
  sanitizeData(data) {
    // Implement sanitization logic as needed
    // For example, redact email addresses or other sensitive fields
    return data;
  },
};

/**
 * Discord Utility Functions
 */
const DiscordUtils = {
  /**
   * Splits content into multiple embeds considering Discord limitations.
   * @param {string} content - The content to be split into embeds.
   * @param {object} options - Options for the embeds.
   * @param {string} options.title - Title for the first embed.
   * @param {object} options.footer - Footer for the embeds.
   * @param {object} options.thumbnail - Thumbnail for the embeds.
   * @param {number} options.color - Color for the embeds.
   * @param {Array} options.fields - Fields to include in the first embed.
   * @returns {Array} - An array of embed objects.
   */
  splitContentIntoEmbeds(content, options = {}) {
    const {
      title = '',
      footer = null,
      thumbnail = null,
      color = null,
      fields = [],
    } = options;

    const MAX_EMBED_DESCRIPTION_LENGTH = CONFIG.MAX_DESCRIPTION_LENGTH;
    const MIN_CHUNK_MERGE_THRESHOLD = CONFIG.MIN_CHUNK_MERGE_THRESHOLD;

    const embeds = [];
    let remainingContent = content;

    // If title or fields are provided, create the first embed with them
    if (title || fields.length > 0) {
      // Create the first embed
      const embed = {
        title: Utils.truncateText(title, 256),
        description: '',
        footer,
        thumbnail,
      };

      if (color !== null) {
        embed.color = color;
      }

      if (fields.length > 0) {
        embed.fields = fields.slice(0, 25); // Max 25 fields
      }

      // Handle the description
      let availableLength = MAX_EMBED_DESCRIPTION_LENGTH;
      if (fields.length > 0) {
        // Estimate fields' length
        const fieldsLength = fields.reduce(
          (sum, field) => sum + field.name.length + field.value.length,
          0
        );
        availableLength -= fieldsLength;
      }

      let description = remainingContent.substring(0, availableLength);
      const lastSpaceIndex = description.lastIndexOf(" ");
      if (lastSpaceIndex > -1 && lastSpaceIndex > description.length * 0.8) {
        description = description.substring(0, lastSpaceIndex);
      }

      embed.description = description;
      embeds.push(embed);

      // Update remaining content
      remainingContent = remainingContent.substring(description.length).trim();
    }

    // Now create additional embeds for remaining content
    while (remainingContent.length > 0) {
      let chunk = remainingContent.substring(0, MAX_EMBED_DESCRIPTION_LENGTH);
      const lastSpaceIndex = chunk.lastIndexOf(" ");
      if (lastSpaceIndex > -1 && lastSpaceIndex > chunk.length * 0.8) {
        chunk = chunk.substring(0, lastSpaceIndex);
      }

      // Check if the remaining content is small enough to append
      const nextChunk = remainingContent.substring(chunk.length).trim();
      if (
        nextChunk.length > 0 &&
        nextChunk.length < MIN_CHUNK_MERGE_THRESHOLD &&
        chunk.length + nextChunk.length <= MAX_EMBED_DESCRIPTION_LENGTH
      ) {
        // Append the next chunk to current chunk
        chunk += ' ' + nextChunk;
        remainingContent = '';
      } else {
        remainingContent = remainingContent.substring(chunk.length).trim();
      }

      const embed = {
        description: chunk,
        footer,
        thumbnail,
      };

      if (color !== null) {
        embed.color = color;
      }

      embeds.push(embed);
    }

    return embeds;
  },

  /**
   * Splits an array of embeds into batches considering Discord's limitations.
   * @param {Array} embeds - An array of embeds to send.
   * @returns {Array} - An array of batches, each batch is an array of embeds.
   */
  splitEmbedsIntoBatches(embeds) {
    const MAX_EMBEDS_PER_MESSAGE = CONFIG.MAX_EMBEDS_PER_MESSAGE;
    const MAX_TOTAL_EMBED_SIZE = CONFIG.MAX_TOTAL_EMBED_SIZE;

    const batches = [];
    let currentBatch = [];
    let currentTotalSize = 0;

    for (const embed of embeds) {
      const embedSize = JSON.stringify(embed).length;

      if (
        currentTotalSize + embedSize > MAX_TOTAL_EMBED_SIZE ||
        currentBatch.length >= MAX_EMBEDS_PER_MESSAGE
      ) {
        batches.push([...currentBatch]);
        currentBatch = [];
        currentTotalSize = 0;
      }

      currentBatch.push(embed);
      currentTotalSize += embedSize;
    }

    if (currentBatch.length > 0) {
      batches.push(currentBatch);
    }

    return batches;
  },
};

/**
 * Logging Functions
 */
const Logger = {
  // Array to collect log entries
  logEntries: [],

  // Add a log entry to the logEntries array
  logToDiscord(level, message) {
    const timestamp = new Date().toISOString();
    this.logEntries.push({ timestamp, level, message });
  },

  /**
   * Function to send log entries as embeds to the log channel with attachment
   */
  async sendLogEmbeds(detailedLog) {
    if (!LOG_CHANNEL_ID || this.logEntries.length === 0) return; // If no log channel is specified or no logs, skip

    try {
      // Prepare the log text content
      const logTextContent = this.logEntries
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

      // Map log levels to colors (customizable)
      const logLevelColors = {
        ERROR: 0xff0000, // Red
        WARN: 0xffa500,  // Orange
        INFO: 0x00ff00,  // Green
      };

      // Build log content per level
      const logContentPerLevel = {};

      for (const entry of this.logEntries) {
        if (!logContentPerLevel[entry.level]) {
          logContentPerLevel[entry.level] = '';
        }
        logContentPerLevel[entry.level] += `**[${entry.level}] ${entry.timestamp}**\n${entry.message}\n\n`;
      }

      // Build embeds per log level
      const embeds = [];
      for (const [level, content] of Object.entries(logContentPerLevel)) {
        const embedsForLevel = DiscordUtils.splitContentIntoEmbeds(content, {
          footer: { text: "Log Entry" },
          color: logLevelColors[level] || null,
        });
        embeds.push(...embedsForLevel);
      }
      // Split embeds into batches
      const embedBatches = DiscordUtils.splitEmbedsIntoBatches(embeds);

      // Send embed batches with attachment
      for (const [index, batch] of embedBatches.entries()) {
        const payload = { embeds: batch };

        // Attach the log file only once
        const attachments = index === 0 ? [logAttachment] : [];

        await DiscordAPI.sendEmbedBatch(LOG_CHANNEL_ID, payload, attachments);
      }
    } catch (error) {
      console.error("Error sending log embeds:", error);
    } finally {
      // Clear log entries after sending
      this.logEntries = [];
    }
  },
};

/**
 * Email Handling Functions
 */
const EmailHandler = {
  // Parse the email using PostalMime
  async parseEmail(event) {
    try {
      Logger.logToDiscord("INFO", "Starting to parse email.");
      const rawEmail = await Utils.streamToArrayBuffer(event.raw);
      const parser = new PostalMime();
      const parsedEmail = await parser.parse(rawEmail);
      Logger.logToDiscord("INFO", `Email parsed successfully.\nSubject: ${parsedEmail.subject || "No Subject"}`);
      return parsedEmail;
    } catch (error) {
      console.error("Error parsing email:", error);
      Logger.logToDiscord("ERROR", `Error parsing email: ${error.message}`);
      throw new Error("Failed to parse email content.");
    }
  },

  // Send an auto-reply email
  async sendAutoReply(event, parsedEmail, errorMsg = null) {
    try {
      const timestamp = new Date().toISOString();
      let subject = parsedEmail.subject || "No Subject";

      // Check if the subject already starts with "Re:" (case insensitive)
      if (!/^Re:/i.test(subject.trim())) {
        subject = `Re: ${subject}`;
      }      
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
        Logger.logToDiscord("WARN", warningMsg);
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
        Logger.logToDiscord("INFO", "Auto-reply sent successfully using event.reply().");
      } catch (replyError) {
        console.warn("event.reply() failed, attempting to send a new email:", replyError);
        Logger.logToDiscord("WARN", `event.reply() failed: ${replyError.message}. Attempting to send a new email.`);
        // If event.reply() fails, send a new email instead
        await event.send(message);
        Logger.logToDiscord("INFO", "Auto-reply sent using event.send().");
      }
    } catch (error) {
      console.error("Error sending auto-reply:", error);
      Logger.logToDiscord("ERROR", `Error sending auto-reply: ${error.message}`);
    }
  },

  // Handle attachments and return their links
  async handleAttachments(attachments) {
    const attachmentLinks = [];

    // Batch attachments according to Discord's limits
    const BATCH_SIZE = CONFIG.ATTACHMENT_BATCH_SIZE;
    const batches = [];

    for (let i = 0; i < attachments.length; i += BATCH_SIZE) {
      batches.push(attachments.slice(i, i + BATCH_SIZE));
    }

    for (const batch of batches) {
      const links = await DiscordAPI.uploadAttachmentsToChannel(ATTACHMENTS_CHANNEL, batch);
      if (links && links.length > 0) {
        attachmentLinks.push(...links);
      }
    }

    Logger.logToDiscord("INFO", `Handled ${attachments.length} attachments.`);
    return attachmentLinks;
  },
};

/**
 * Discord API Interaction Functions
 */
const DiscordAPI = {
  // Fetch guild data from Discord API
  async fetchGuildData() {
    try {
      const url = `${CONFIG.DISCORD_API_URL}/guilds/${GUILD_ID}`;
      const response = await Utils.fetchWithRateLimit(url, {
        method: "GET",
        headers: myHeaders,
      });

      if (!response.ok) {
        throw new Error(`Failed to fetch guild data. Status: ${response.status}`);
      }

      const guildData = await response.json();
      Logger.logToDiscord("INFO", `Fetched guild data: ${guildData.name}`);
      return guildData;
    } catch (error) {
      console.error("Error fetching guild data:", error);
      Logger.logToDiscord("ERROR", `Error fetching guild data: ${error.message}`);
      throw new Error("Failed to fetch guild information.");
    }
  },

  // Get the guild icon URL
  getGuildIconURL(guildData) {
    if (guildData.icon) {
      const fileExtension = guildData.icon.startsWith("a_") ? ".gif" : ".png";
      return `https://cdn.discordapp.com/icons/${guildData.id}/${guildData.icon}${fileExtension}`;
    }
    return null;
  },

  // Find Discord member by username
  async findDiscordMember(username) {
    try {
      const fetchMemberURL = `${CONFIG.DISCORD_API_URL}/guilds/${GUILD_ID}/members/search?query=${encodeURIComponent(
        username
      )}&limit=1000`;

      const response = await Utils.fetchWithRateLimit(fetchMemberURL, {
        method: "GET",
        headers: myHeaders,
      });
      if (!response.ok) {
        throw new Error(`Failed to fetch Discord members. Status: ${response.status}`);
      }

      const members = await response.json();
      const member = members.find((mem) => mem.user.username === username);
      if (member) {
        Logger.logToDiscord("INFO", `Found Discord member: ${member.user.username}#${member.user.discriminator}`);
      } else {
        Logger.logToDiscord("WARN", `Discord member not found for username: ${username}`);
      }
      return member;
    } catch (error) {
      console.error("Error fetching Discord member:", error);
      Logger.logToDiscord("ERROR", `Error fetching Discord member: ${error.message}`);
      throw new Error(`Failed to find Discord member with username: ${username}`);
    }
  },

  // Check if the member has the required role(s)
  hasRequiredRoles(member) {
    if (!member) return false;
    const memberRoles = member.roles; // Array of role IDs the member has

    if (ROLES_REQUIRED.length === 0) return true; // No roles required

    const hasRoles = ROLES_REQUIRED.some((role) => memberRoles.includes(role));
    return hasRoles;
  },

  // Create a DM channel with the user
  async createDmChannel(recipientId) {
    try {
      const createDmURL = `${CONFIG.DISCORD_API_URL}/users/@me/channels`;
      const response = await Utils.fetchWithRateLimit(createDmURL, {
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
      Logger.logToDiscord("INFO", `Created DM channel with user ID: ${recipientId}`);
      return dmChannel;
    } catch (error) {
      console.error("Error creating DM channel:", error);
      Logger.logToDiscord("ERROR", `Error creating DM channel: ${error.message}`);
      throw new Error("Failed to create a direct message channel with the member.");
    }
  },

  // Send a batch of embeds to a Discord channel
  async sendEmbedBatch(channelId, payload, attachments = null) {
    const options = {
      method: "POST",
      headers: {
        Authorization: `Bot ${TOKEN}`,
      },
    };

    if (attachments && attachments.length > 0) {
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

    const response = await Utils.fetchWithRateLimit(
      `${CONFIG.DISCORD_API_URL}/channels/${channelId}/messages`,
      options
    );

    if (!response.ok) {
      const errorText = await response.text();
      Logger.logToDiscord("ERROR", `Failed to send embed batch: ${errorText}`);
      throw new Error(`Failed to send message. Status: ${response.status}. Error: ${errorText}`);
    }

    return await response.json();
  },

  // Send multiple embeds with batching
  async sendEmbedsWithBatching(channelId, embeds, attachments = null) {
    // Now uses DiscordUtils.splitEmbedsIntoBatches
    const batches = DiscordUtils.splitEmbedsIntoBatches(embeds);

    for (const [index, batch] of batches.entries()) {
      const payload = { embeds: batch };
      const attach = index === batches.length - 1 && attachments ? attachments : [];
      await this.sendEmbedBatch(channelId, payload, attach);
    }

    Logger.logToDiscord("INFO", `Sent ${embeds.length} embeds to channel ID: ${channelId}`);
    return batches;
  },

  // Send the text content as an attachment
  async sendTextAttachment(channelId, textContent, attachmentLinks) {
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

      const response = await Utils.fetchWithRateLimit(
        `${CONFIG.DISCORD_API_URL}/channels/${channelId}/messages`,
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
        Logger.logToDiscord("ERROR", `Failed to send text attachment: ${errorText}`);
        throw new Error(
          `Failed to send message with attachment. Status: ${response.status}. Error: ${errorText}`
        );
      }

      Logger.logToDiscord("INFO", `Sent text attachment to channel ID: ${channelId}`);
      return await response.json();
    } catch (error) {
      console.error("Error sending text attachment:", error);
      Logger.logToDiscord("ERROR", `Error sending text attachment: ${error.message}`);
      throw new Error("Failed to send the text attachment.");
    }
  },

  // Upload attachments to a Discord channel and return their URLs
  async uploadAttachmentsToChannel(channelId, attachments) {
    try {
      const formData = new FormData();

      // Append attachments (up to Discord's limit per message)
      attachments.forEach((attachment, index) => {
        const blob = new Blob([attachment.content], { type: attachment.contentType });
        formData.append(`files[${index}]`, blob, attachment.filename);
      });

      const sendMessageURL = `${CONFIG.DISCORD_API_URL}/channels/${channelId}/messages`;

      const response = await Utils.fetchWithRateLimit(sendMessageURL, {
        method: "POST",
        headers: {
          Authorization: `Bot ${TOKEN}`,
        },
        body: formData,
      });

      if (!response.ok) {
        const errorText = await response.text();
        Logger.logToDiscord("ERROR", `Failed to upload attachments: ${errorText}`);
        throw new Error(
          `Failed to upload attachments. Status: ${response.status}. Error: ${errorText}`
        );
      }

      const messageData = await response.json();

      // Extract attachment URLs
      if (messageData.attachments && messageData.attachments.length > 0) {
        const urls = messageData.attachments.map((attachment) => attachment.url);
        Logger.logToDiscord("INFO", `Uploaded ${urls.length} attachments to channel ID: ${channelId}`);
        return urls;
      } else {
        const errorMsg = "No attachment URLs found in message data.";
        console.error(errorMsg);
        Logger.logToDiscord("WARN", errorMsg);
        return [];
      }
    } catch (error) {
      console.error("Error uploading attachments:", error);
      Logger.logToDiscord("ERROR", `Error uploading attachments: ${error.message}`);
      return [];
    }
  },

  // Send the email content as embeds to Discord
  async sendEmbedMessage(channelId, parsedEmail, event, attachmentLinks) {
    try {
      // Fetch guild data and icon
      const guildData = await this.fetchGuildData();
      const thumbnailUrl = this.getGuildIconURL(guildData);

      // Prepare email fields with fallback
      const fromField = parsedEmail.from?.address || parsedEmail.from?.[0]?.address || event.from || "Unknown";
      const toField = parsedEmail.to?.map((addr) => addr.address).join(", ") || event.to || "Unknown";

      // Extract email text content
      let emailTextContent = parsedEmail.text || Utils.extractTextFromHtml(parsedEmail.html);
      emailTextContent = emailTextContent || "(No text content)";

      // Sanitize email text content for privacy
      emailTextContent = Utils.sanitizeData(emailTextContent);

      // Construct the initial embed options
      const title = `📧 ${parsedEmail.subject || "New Email Received"}`;
      const footerText = "📬 Sent via Clanflare Email System";
      const timestamp = new Date().toISOString();

      const fields = [
        { name: "👤 From", value: fromField },
        { name: "📩 To", value: toField },
        { name: "📅 Date", value: parsedEmail.date || timestamp },
        { name: "📎 Attachments", value: `${attachmentLinks.length}` },
      ];

      const embedOptions = {
        title,
        footer: { text: footerText },
        thumbnail: { url: thumbnailUrl },
        fields,
      };

      // Use DiscordUtils to split email content into embeds
      const embeds = DiscordUtils.splitContentIntoEmbeds(emailTextContent, embedOptions);

      // If there are attachment links, add them as additional embeds
      if (attachmentLinks && attachmentLinks.length > 0) {
        const attachmentContent = attachmentLinks
          .map((link, index) => `[Attachment ${index + 1}](${link})`)
          .join("\n");

        const attachmentEmbeds = DiscordUtils.splitContentIntoEmbeds(attachmentContent, {
          footer: { text: footerText },
          thumbnail: { url: thumbnailUrl },
        });

        embeds.push(...attachmentEmbeds);
      }

      // Send embeds with batching
      await this.sendEmbedsWithBatching(channelId, embeds);

      // Send the text content as an attachment
      if (emailTextContent && emailTextContent.length > 0) {
        await this.sendTextAttachment(channelId, emailTextContent, attachmentLinks);
      }

      Logger.logToDiscord("INFO", `Sent email content to channel ID: ${channelId}`);
    } catch (error) {
      console.error("Error sending embed message:", error);
      Logger.logToDiscord("ERROR", `Error sending embed message: ${error.message}`);
      throw new Error("Failed to send an embed message to the member.");
    }
  },
};

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
      Logger.logToDiscord("INFO", "Email event received.");

      // Log event details
      Logger.logToDiscord("INFO", `Event details:\nFrom: ${event.from}\nTo: ${event.to}`);

      parsedEmail = await EmailHandler.parseEmail(event);
      const username = parsedEmail.to[0]?.address.split("@")[0] || "Unknown";
      Logger.logToDiscord("INFO", `Parsed email intended for username: ${username}`);

      // Log email details
      Logger.logToDiscord(
        "INFO",
        `Email details:\nSubject: ${parsedEmail.subject || "No Subject"}\nFrom: ${parsedEmail.from?.address || parsedEmail.from?.[0]?.address || event.from || "Unknown"
        }\nTo: ${parsedEmail.to?.map((addr) => addr.address).join(", ") || event.to || "Unknown"
        }\nDate: ${parsedEmail.date || new Date().toISOString()
        }\nAttachments: ${parsedEmail.attachments?.length || 0}`
      );

      // Sanitize and log parsedEmail data
      const sanitizedParsedEmail = Utils.sanitizeData(parsedEmail);
      Logger.logToDiscord("INFO", `Parsed Email Data:\n${JSON.stringify(sanitizedParsedEmail, null, 2)}`);

      // Log ctx data
      Logger.logToDiscord("INFO", `Context Data:\n${JSON.stringify(ctx, null, 2)}`);

      // Handle attachments
      let attachmentLinks = [];
      if (parsedEmail.attachments && parsedEmail.attachments.length > 0) {
        attachmentLinks = await EmailHandler.handleAttachments(parsedEmail.attachments);
      }

      // Determine target channel
      let targetChannelId;
      if (CHANNEL_MAP[username]) {
        targetChannelId = CHANNEL_MAP[username];
        Logger.logToDiscord("INFO", `Using channel from CHANNEL_MAP: ${targetChannelId}`);
      } else {
        const member = await DiscordAPI.findDiscordMember(username);

        if (member) {
          if (!DiscordAPI.hasRequiredRoles(member)) {
            const errorMsg = "Member does not have the required role(s).";
            Logger.logToDiscord("WARN", errorMsg);
            throw new Error(errorMsg);
          }
          const dm = await DiscordAPI.createDmChannel(member.user.id);
          targetChannelId = dm.id;
          Logger.logToDiscord("INFO", `DM channel created: ${targetChannelId}`);
        } else if (CHANNEL_MAP["others"]) {
          targetChannelId = CHANNEL_MAP["others"];
          Logger.logToDiscord("INFO", `Using 'others' channel from CHANNEL_MAP: ${targetChannelId}`);
        } else {
          const errorMsg = "No target channel found for the recipient.";
          Logger.logToDiscord("ERROR", errorMsg);
          throw new Error(errorMsg);
        }
      }

      // Send message to Discord
      await DiscordAPI.sendEmbedMessage(targetChannelId, parsedEmail, event, attachmentLinks);

      // Send auto-reply
      await EmailHandler.sendAutoReply(event, parsedEmail);

      Logger.logToDiscord("INFO", "Email processing completed successfully.");
    } catch (error) {
      console.error("Error handling email event:", error);
      Logger.logToDiscord("ERROR", `Error handling email event: ${error.message}`);
      if (parsedEmail) {
        await EmailHandler.sendAutoReply(event, parsedEmail, error.message);
      }
    } finally {
      // Prepare detailed log data
      const detailedLog = {
        parsedEmail: Utils.sanitizeData(parsedEmail) || {},
        event: {
          from: event.from,
          to: event.to,
          headers: event.headers,
        },
        ctx: ctx || {},
      };

      // Send all accumulated logs as embeds with attachment
      await Logger.sendLogEmbeds(detailedLog);
    }
  },
};
