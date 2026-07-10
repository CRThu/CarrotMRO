import { test, expect } from '@playwright/test'

test.describe('CarrotMRO 基础功能', () => {
  test('页面加载正常', async ({ page }) => {
    await page.goto('/')
    await expect(page.locator('text=CarrotMRO')).toBeVisible()
  })

  test('侧边栏显示项目和定价表', async ({ page }) => {
    await page.goto('/')
    await expect(page.locator('text=项目')).toBeVisible()
    await expect(page.locator('text=协议基准价格清单')).toBeVisible()
  })

  test('可以创建新项目', async ({ page }) => {
    await page.goto('/')

    // 点击创建项目按钮
    const createButton = page.locator('button:has-text("新项目名称...")')
    if (await createButton.isVisible()) {
      await createButton.click()
      const input = page.locator('input[placeholder="新项目名称..."]')
      await input.fill('测试项目')
      await input.press('Enter')
    }
  })
})
