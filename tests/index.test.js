import { describe, it, expect, beforeEach, vi } from 'vitest';
import { JSDOM } from 'jsdom';
import fs from 'fs';
import path from 'path';

const htmlSource = fs.readFileSync(path.resolve(__dirname, '../index.html'), 'utf8');

describe('Move-All Frontend index.html Unit Tests', () => {
  let dom;
  let window;
  let document;

  beforeEach(() => {
    dom = new JSDOM(htmlSource, {
      url: 'http://localhost/',
      runScripts: 'dangerously',
      resources: 'usable'
    });
    window = dom.window;
    document = window.document;

    // Mock window methods for test environment
    window.fetch = vi.fn().mockImplementation(() =>
      Promise.resolve({
        json: () => Promise.resolve({
          success: true,
          isAdmin: true,
          leaderboard: [
            { email: 'zach.yang@north.com.tw', name: '🥷 忍者哈特利', validDays: 1, sportCount: 1, totalSteps: 8000, points: 2, tier: '尚未達標', reward: 0 }
          ],
          teams: [
            { teamName: '諾思隊', captainEmail: 'zach.yang@north.com.tw', members: ['zach.yang@north.com.tw'], teamSteps: 8000, teamPoints: 3, statusBadge: '尚未達標', rewardText: '參加獎' }
          ],
          records: [
            { rowId: 1, email: 'zach.yang@north.com.tw', name: '🥷 忍者哈特利', date: '2026/09/04', steps: 8000, status: '通過', approvedSteps: 8000 }
          ],
          chatMessages: [
            { id: 'chat_1', email: 'zach.yang@north.com.tw', name: '🥷 忍者哈特利', message: 'Hello world', time: '09/04 12:00' }
          ]
        })
      })
    );
    window.alert = vi.fn();

    // Find all <script> blocks without src and evaluate them in window context
    const regex = /<script(?![^>]*src=)([^>]*)>([\s\S]*?)<\/script>/gi;
    let match;
    while ((match = regex.exec(htmlSource)) !== null) {
      const code = match[2];
      if (code && code.trim()) {
        try {
          window.eval(code);
        } catch (err) {
          console.error("Eval script error:", err);
        }
      }
    }
  });

  it('1. should initialize window and DOM elements properly', () => {
    expect(document.getElementById('loginGate')).not.toBeNull();
    expect(document.getElementById('mainApp')).not.toBeNull();
    expect(document.getElementById('googleLoginBtnGate')).not.toBeNull();
  });

  it('2. should trigger showLoggedOutUI when no user is saved', () => {
    window.showLoggedOutUI();
    const loginGate = document.getElementById('loginGate');
    const mainApp = document.getElementById('mainApp');
    expect(loginGate.classList.contains('hidden')).toBe(false);
    expect(mainApp.classList.contains('hidden')).toBe(true);
  });

  it('3. should trigger showLoggedInUI when user is logged in', () => {
    window.currentEmail = 'zach.yang@north.com.tw';
    window.showLoggedInUI('🥷 忍者哈特利');
    const loginGate = document.getElementById('loginGate');
    const mainApp = document.getElementById('mainApp');
    expect(loginGate.classList.contains('hidden')).toBe(true);
    expect(mainApp.classList.contains('hidden')).toBe(false);
  });

  it('4. should handle handleCredentialResponse correctly on Google login', () => {
    const mockPayload = {
      email: 'zach.yang@north.com.tw',
      name: '忍者哈特利'
    };
    window.handleCredentialResponse(mockPayload);
    expect(window.currentEmail).toBe('zach.yang@north.com.tw');
    expect(window.localStorage.getItem('north_user')).toContain('zach.yang@north.com.tw');
    
    const loginGate = document.getElementById('loginGate');
    expect(loginGate.classList.contains('hidden')).toBe(true);
  });

  it('5. should call promptFallbackEmailLogin function directly (no UI button)', () => {
    window.prompt = () => 'zach.yang@north.com.tw';
    window.promptFallbackEmailLogin();
    expect(window.currentEmail).toBe('zach.yang@north.com.tw');
    expect(window.localStorage.getItem('north_user')).toContain('zach.yang@north.com.tw');
  });

  it('6. should render individual leaderboard correctly without error', () => {
    const mockList = [
      { email: 'zach.yang@north.com.tw', name: '忍者', validDays: 1, sportCount: 1, totalSteps: 8000, points: 2, tier: '尚未達標', reward: 0 }
    ];
    window.renderLeaderboard(mockList);
    const tbody = document.getElementById('leaderboardBody');
    expect(tbody.innerHTML).toContain('zach.yang@north.com.tw');
    expect(tbody.innerHTML).toContain('2 點');
  });

  it('7. should render team leaderboard correctly without error', () => {
    const mockTeams = [
      { teamName: '諾思隊', captainEmail: 'zach.yang@north.com.tw', members: ['zach.yang@north.com.tw'], teamSteps: 8000, teamPoints: 3, statusBadge: '尚未達標', rewardText: '參加獎' }
    ];
    window.renderTeamLeaderboard([], mockTeams);
    const tbody = document.getElementById('teamLeaderboardBody');
    expect(tbody.innerHTML).toContain('諾思隊');
    expect(tbody.innerHTML).toContain('3 點');
  });

  it('8. should render feed and chat messages without error', () => {
    const mockRecords = [
      { rowId: 10, email: 'zach.yang@north.com.tw', name: '忍者', date: '2026/09/04', steps: 8000, status: '通過' }
    ];
    const mockChat = [
      { id: 'chat_10', email: 'zach.yang@north.com.tw', name: '忍者', message: 'Hello test', time: '12:00' }
    ];
    window.renderFeed(mockRecords);
    window.renderChatMessages(mockChat);

    const feedList = document.getElementById('feedList');
    const chatList = document.getElementById('chatMessageList');
    expect(feedList.innerHTML).toContain('8,000');
    expect(chatList.innerHTML).toContain('Hello test');
  });

  it('9. should handle user logoutUser cleanly', () => {
    window.currentEmail = 'zach.yang@north.com.tw';
    window.logoutUser();
    expect(window.currentEmail).toBe('');
    expect(window.localStorage.getItem('north_user')).toBeNull();
    const loginGate = document.getElementById('loginGate');
    expect(loginGate.classList.contains('hidden')).toBe(false);
  });
});
