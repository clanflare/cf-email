import { EmailMessage } from "cloudflare:email";
import { createMimeMessage } from "mimetext";
const PostalMime = require("postal-mime");

const API_VERSION = 10;
const DISCORD_API_URL = `https://discord.com/api/v${API_VERSION}`;
const GUILD_ID = "1209956399073992745";
const TOKEN = "";

const myHeaders = new Headers();
myHeaders.append("Authorization", `Bot ${TOKEN}`);
myHeaders.append("Content-Type", "application/json");

const requestOptions = {
  headers: myHeaders,
  redirect: "follow"
};

// Utility function to convert a stream to an ArrayBuffer
async function streamToArrayBuffer(stream, streamSize) {
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
}

// Function to parse the email
async function parseEmail(event) {
  try {
    const rawEmail = await streamToArrayBuffer(event.raw, event.rawSize);
    const parser = new PostalMime.default();
    return await parser.parse(rawEmail);
  } catch (error) {
    throw new Error(`Error parsing email: ${error.message}`);
  }
}

// Function to find Discord member by username
async function findDiscordMember(username) {
  try {
    const fetchMemberURL = `${DISCORD_API_URL}/guilds/${GUILD_ID}/members/search?query=${username}&limit=1000`;
    const response = await fetch(fetchMemberURL, requestOptions);
    const members = await response.json();
    return members.find(mem => mem.user.username === username);
  } catch (error) {
    throw new Error(`Error fetching Discord member: ${error.message}`);
  }
}

// Function to create a DM channel with the user
async function createDmChannel(recipientId) {
  try {
    const createDmURL = `${DISCORD_API_URL}/users/@me/channels`;
    const response = await fetch(createDmURL, {
      method: "POST",
      body: JSON.stringify({ recipient_id: recipientId }),
      ...requestOptions
    });
    return await response.json();
  } catch (error) {
    throw new Error(`Error creating DM channel: ${error.message}`);
  }
}

// Function to send a message in a Discord channel
async function sendMessage(channelId, content) {
  try {
    const sendMessageURL = `${DISCORD_API_URL}/channels/${channelId}/messages`;
    const response = await fetch(sendMessageURL, {
      method: "POST",
      body: JSON.stringify({ content }),
      ...requestOptions
    });
    return await response.json();
  } catch (error) {
    throw new Error(`Error sending message: ${error.message}`);
  }
}

// Function to send an auto-reply email, optionally with an error message
async function sendAutoReply(event, parsedEmail, errorMsg = null) {
  try { 
    const messageData = errorMsg
      ? `An error occurred while processing your email. Error details: ${errorMsg}`
      : `This is an automated reply to your email with the subject ${parsedEmail.subject}.\nYour email has been succesfully delivered to ${parsedEmail.to[0].address.split("@")[0]}`;

    const msg = createMimeMessage();
    msg.setSender({ name: "Auto-replier", addr: event.to });
    msg.setRecipient(event.from);
    msg.setSubject(`Re: ${parsedEmail.subject}`);
    if (parsedEmail?.messageId) {
      msg.setHeader("In-Reply-To", parsedEmail.messageId);
    }
    msg.addMessage({
      contentType: "text/plain",
      data: messageData,
    });

    const message = new EmailMessage(event.to, event.from, msg.asRaw());
    await event.reply(message);
  } catch (error) {
    console.error("Error sending auto-reply:", error);
  }
}

// Main function to handle the email event
export default {
  async email(event, env, ctx) {
    let parsedEmail = null;
    try {
      parsedEmail = await parseEmail(event);
      const username = parsedEmail.to[0].address.split("@")[0];
      const member = await findDiscordMember(username);

      if (member) {
        const dm = await createDmChannel(member.user.id);
        const dmContent = `## From: ${event.from}\n## Subject: ${parsedEmail.subject}\n ${parsedEmail.text}`
        await sendMessage(dm.id, dmContent);
      } else {
        throw new Error(`No member found with username: ${username}`);
      }

      await sendAutoReply(event, parsedEmail);
    } catch (error) {
      console.error("Error handling email event:", error);
      await sendAutoReply(event, parsedEmail, error.message);
    }
  },
};