import i18next from 'i18next'
import { Docx, docx, Toast } from '@dolphin/lark'
import { generatePublicUrl, makePublicUrlEffective } from '@dolphin/lark/image'
import { isDefined } from '@dolphin/common'
import { CommonTranslationKey, en, Namespace, zh } from '../common/i18n'
import { confirm } from '../common/notification'
import { reportBug } from '../common/issue'
import {
  transformMentionUsers,
  transformTableWithParents,
} from '../common/utils'
import {
  getSettings,
  SettingKey,
  TableWithNonPhrasingContent,
  Grid,
} from '../common/settings'

const enum TranslationKey {
  UNKNOWN_ERROR = 'unknown_error',
  CONTENT_LOADING = 'content_loading',
  NOT_SUPPORT = 'not_support',
  NOT_SUPPORT_DOC_1_0 = 'not_support_doc_1_0',
  PROCESSING_IMAGES = 'processing_images',
}

i18next
  .init({
    lng: docx.language,
    resources: {
      en: {
        translation: {
          [TranslationKey.UNKNOWN_ERROR]: 'Unknown error during download',
          [TranslationKey.CONTENT_LOADING]:
            'Part of the content is still loading and cannot be copied at the moment. Please wait for loading to complete and retry',
          [TranslationKey.NOT_SUPPORT]:
            'This is not a lark document page and cannot be copied as Markdown',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            'This is a old version lark document page and cannot be copied as Markdown',
          [TranslationKey.PROCESSING_IMAGES]:
            'Processing images, please wait...',
        },
        ...en,
      },
      zh: {
        translation: {
          [TranslationKey.UNKNOWN_ERROR]: '下载过程中出现未知错误',
          [TranslationKey.CONTENT_LOADING]:
            '部分内容仍在加载中，暂时无法复制。请等待加载完成后重试',
          [TranslationKey.NOT_SUPPORT]:
            '这不是一个飞书文档页面，无法复制为 Markdown',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            '这是一个旧版飞书文档页面,无法复制为 Markdown',
          [TranslationKey.PROCESSING_IMAGES]:
            '正在处理图片，请稍候...',
        },
        ...zh,
      },
    },
  })
  .catch(console.error)

// 使用 Chrome 扩展特权下载图片并转换为 base64
const downloadImageAsBase64 = async (url: string): Promise<string> => {
  try {
    // 先尝试从页面的 img 元素获取（如果是 blob URL）
    if (url.startsWith('blob:')) {
      const img = Array.from(document.querySelectorAll('img')).find(
        el => el.src === url
      ) as HTMLImageElement

      if (img?.complete) {
        const canvas = document.createElement('canvas')
        canvas.width = img.naturalWidth || img.width
        canvas.height = img.naturalHeight || img.height
        const ctx = canvas.getContext('2d')
        if (ctx && canvas.width > 0 && canvas.height > 0) {
          ctx.drawImage(img, 0, 0)
          const base64 = canvas.toDataURL('image/png')
          // 检查是否成功转换（base64 应该很长）
          if (base64.length > 100) {
            return base64
          }
        }
      }
    }

    // 对于非 blob URL，使用 fetch 下载
    const response = await fetch(url, { credentials: 'include' })
    const blob = await response.blob()

    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onloadend = () => resolve(reader.result as string)
      reader.onerror = reject
      reader.readAsDataURL(blob)
    })
  } catch (error) {
    console.error('Failed to download image:', url, error)
    // Fallback: 尝试从页面已加载的图片元素提取
    try {
      const img = Array.from(document.querySelectorAll('img')).find(
        el => el.src === url || el.currentSrc === url
      ) as HTMLImageElement

      if (img?.complete && img.naturalWidth > 0) {
        const canvas = document.createElement('canvas')
        canvas.width = img.naturalWidth
        canvas.height = img.naturalHeight
        const ctx = canvas.getContext('2d')
        if (ctx) {
          ctx.drawImage(img, 0, 0)
          const base64 = canvas.toDataURL('image/png')
          if (base64.length > 100) return base64
        }
      }
    } catch (fallbackError) {
      console.error('Fallback extraction failed:', fallbackError)
    }
    return url
  }
}

const main = async () => {
  if (docx.isDoc) {
    Toast.warning({ content: i18next.t(TranslationKey.NOT_SUPPORT_DOC_1_0) })

    return
  }

  if (!docx.isDocx) {
    Toast.warning({ content: i18next.t(TranslationKey.NOT_SUPPORT) })

    return
  }

  if (!docx.isReady()) {
    Toast.warning({
      content: i18next.t(TranslationKey.CONTENT_LOADING),
    })

    return
  }

  const settings = await getSettings([
    SettingKey.TableWithNonPhrasingContent,
    SettingKey.Grid,
    SettingKey.TextHighlight,
  ])

  const { root, images, tableWithParents, mentionUsers } = docx.intoMarkdownAST(
    {
      highlight: settings[SettingKey.TextHighlight],
      flatGrid: settings[SettingKey.Grid] === Grid.Flatten,
    },
  )

  console.log('Total images:', images.length)

  await transformMentionUsers(mentionUsers)

  // 收集所有有 token 的图片
  const tokens = images
    .map(image => {
      if (!image.data?.token) return null

      const { token } = image.data
      const publicUrl = generatePublicUrl(token)
      const code = new URL(publicUrl).searchParams.get('code')
      if (!code) return null

      image.url = publicUrl

      return [token, code]
    })
    .filter(isDefined)

  transformTableWithParents(tableWithParents, {
    transformGridToHtml: settings[SettingKey.Grid] === Grid.ToHTML,
    transformInvalidTablesToHtml:
      settings[SettingKey.TableWithNonPhrasingContent] ===
      TableWithNonPhrasingContent.ToHTML,
  })

  const markdown = Docx.stringify(root)

  if (!window.document.hasFocus()) {
    const confirmed = await confirm()
    if (!confirmed) {
      return
    }
  }

  const writeToClipboard = (
    Object.getPrototypeOf(window.navigator.clipboard) as Clipboard
  ).write.bind(window.navigator.clipboard)

  await writeToClipboard([
    new ClipboardItem({
      'text/plain': new Blob([markdown], { type: 'text/plain' }),
    }),
  ])

  // 尝试激活公开 URL，如果失败则转为 base64
  if (tokens.length > 0) {
    const isSuccess = await makePublicUrlEffective(
      Object.fromEntries(tokens) as Record<string, string>,
    )

    if (!isSuccess) {
      // URL 激活失败，说明没有编辑权限，转换为 base64
      await Promise.all(
        images.map(async image => {
          if (!image.data?.fetchSources) return
          const sources = await image.data.fetchSources()

          if (sources) {
            const originalUrl = sources.originSrc || sources.src

            // 无论是 blob URL 还是 CDN URL，都转换为 base64
            if (originalUrl) {
              image.url = await downloadImageAsBase64(originalUrl)
            }
          }
        }),
      )

      // 重新生成 markdown 并复制
      const newMarkdown = Docx.stringify(root)
      await writeToClipboard([
        new ClipboardItem({
          'text/plain': new Blob([newMarkdown], { type: 'text/plain' }),
        }),
      ])
    }
  }
}

main().catch((error: unknown) => {
  Toast.error({
    content: i18next.t(TranslationKey.UNKNOWN_ERROR),
    actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
      ns: Namespace.COMMON,
    }),
    onActionClick: () => {
      reportBug(error)
    },
  })
})
