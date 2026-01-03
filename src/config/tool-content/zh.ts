/**
 * Chinese Tool Content
 * Requirements: 3.1 - Multi-language support
 * 
 * 中文工具内容 - 包含所有67个PDF工具的详细描述、使用说明、用例和常见问题
 */

import type { ToolContent } from '@/types/tool';

export const toolContentZh: Record<string, ToolContent> = {
  // ==================== 热门工具 ====================
  'pdf-multi-tool': {
    title: 'PDF多功能工具',
    metaDescription: '一站式PDF编辑器：合并、拆分、整理、删除、旋转和提取页面，功能强大。',
    keywords: ['pdf多功能工具', 'pdf编辑器', '合并pdf', '拆分pdf', '整理pdf', '一站式pdf'],
    description: `
      <p>PDF多功能工具是您处理所有PDF页面管理任务的综合解决方案。这款强大的一站式工具将多种PDF操作整合到一个直观的界面中，为您节省时间和精力。</p>
      <p>无论您需要合并多个文档、将大型PDF拆分成小文件、重新整理页面、删除不需要的内容、旋转页面还是提取特定部分，这个工具都能轻松处理，无需在不同应用程序之间切换。</p>
      <p>所有处理都直接在您的浏览器中完成，确保您的文档保持私密和安全。不会将任何文件上传到服务器。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '将PDF文件拖放到上传区域，或点击浏览并从设备中选择文件。' },
      { step: 2, title: '选择操作', description: '从可用操作中选择：合并、拆分、整理、删除页面、旋转、添加空白页或提取页面。' },
      { step: 3, title: '配置选项', description: '根据所选操作调整设置，如页面范围、旋转角度或合并顺序。' },
      { step: 4, title: '处理并下载', description: '点击处理按钮，操作完成后下载修改后的PDF。' },
    ],
    useCases: [
      { title: '文档准备', description: '通过删除不必要的页面、重新排序内容和合并多个文件来准备提交文档。', icon: 'file-check' },
      { title: '报告汇编', description: '合并多个报告部分，添加封面页，将章节整理成一个专业文档。', icon: 'book-open' },
      { title: '档案管理', description: '将大型档案文件拆分成可管理的部分，提取相关页面，重新整理历史文档。', icon: 'archive' },
    ],
    faq: [
      { question: '一次可以处理多少个PDF？', answer: '您可以同时上传和处理最多10个PDF文件，合并最大大小为500MB。' },
      { question: '书签会被保留吗？', answer: '是的，合并PDF时，工具会保留现有书签，并可选择将它们合并成统一的书签结构。' },
      { question: '有页数限制吗？', answer: '没有严格的页数限制。该工具可以处理数百页的文档，但非常大的文件可能需要更长的处理时间。' },
    ],
  },

  'merge-pdf': {
    title: '合并PDF',
    metaDescription: '将多个PDF文件合并成一个文档。免费在线PDF合并器，支持拖放重新排序。',
    keywords: ['合并pdf', '组合pdf', '连接pdf', 'pdf合并器', '拼接pdf'],
    description: `
      <p>合并PDF允许您快速轻松地将多个PDF文档合并成一个文件。无论您是整合报告、合并扫描文档还是组装演示文稿，这个工具都能使过程变得无缝。</p>
      <p>只需上传文件，使用拖放功能按所需顺序排列，然后将它们合并成一个连贯的文档。该工具保留原始文件的质量，并可选择保留每个源文档的书签。</p>
      <p>所有合并操作都在您的浏览器本地进行，确保敏感文档的完全隐私。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '将多个PDF文件拖放到上传区域，或点击从设备中选择文件。' },
      { step: 2, title: '排列顺序', description: '拖放文件缩略图以按所需顺序排列。' },
      { step: 3, title: '合并并下载', description: '点击合并按钮组合所有文件，然后下载合并后的PDF。' },
    ],
    useCases: [
      { title: '合并报告', description: '将月度或季度报告合并成一个年度文档，便于分发和存档。', icon: 'file-text' },
      { title: '组装作品集', description: '将多个项目文档、证书或工作样本合并成专业作品集。', icon: 'briefcase' },
      { title: '整合发票', description: '将多张发票或收据合并成一个文档，用于会计和记录保存。', icon: 'receipt' },
    ],
    faq: [
      { question: '可以合并多少个PDF？', answer: '您可以一次合并最多100个PDF文件，总大小最高可达500MB。' },
      { question: '合并后的PDF会保持原始质量吗？', answer: '是的，合并过程保留所有文档的原始质量，不会进行任何压缩或质量损失。' },
      { question: '可以合并受密码保护的PDF吗？', answer: '受密码保护的PDF需要先解密。请使用我们的解密PDF工具在合并前移除密码。' },
    ],
  },

  'rotate-custom': {
    title: '自定义旋转 PDF',
    metaDescription: '按任意角度旋转PDF页面。精确的自定义旋转，用于校正扫描文档。',
    keywords: ['自定义旋转pdf', 'pdf旋转任意角度', '校正pdf', 'pdf歪斜校正'],
    description: `
      <p>自定义旋转 PDF 工具让您可以精确控制PDF页面的方向。与仅支持90度增量的标准旋转工具不同，此工具允许您按任何特定角度旋转页面。</p>
      <p>非常适合校正扫描时稍微倾斜的文档，或将图表和图纸调整到正确的方向。您可以校正单个页面或对整个文档应用相同的旋转。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档在实现完美对齐的同时保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '上传包含需要旋转页面的PDF文件。' },
      { step: 2, title: '设置旋转角度', description: '为每个页面输入精确的旋转度数，或为所有页面设置批量角度。' },
      { step: 3, title: '预览和调整', description: '使用实时预览确保页面完美对齐。' },
      { step: 4, title: '应用并下载', description: '点击旋转应用更改并下载校正后的PDF。' },
    ],
    useCases: [
      { title: '扫描文档', description: '校正扫描时进纸倾斜的页面。', icon: 'scan' },
      { title: '技术图纸', description: '精确调整技术图表和平面图的方向。', icon: 'ruler' },
      { title: '创意排版', description: '通过将页面旋转到特定艺术角度来创建独特的布局。', icon: 'pen-tool' },
    ],
    faq: [
      { question: '可以按小数旋转吗，例如45.5度？', answer: '目前工具支持整数度数，但我们正在努力启用小数精度。' },
      { question: '这会影响页面内容吗？', answer: '内容会被视觉旋转。页面大小会自动调整以适应旋转后的内容。' },
      { question: '可以只旋转一个页面吗？', answer: '是的，您可以为任何单个页面设置自定义旋转角度，同时保持其他页面不变。' },
    ],
  },

  'split-pdf': {
    title: '拆分PDF',
    metaDescription: '将PDF文件拆分成多个文档。提取特定页面或按页面范围分割。',
    keywords: ['拆分pdf', '分割pdf', '分离pdf', '提取页面', 'pdf拆分器'],
    description: `
      <p>拆分PDF使您能够将单个PDF文档分成多个较小的文件。这非常适合提取特定章节、分离合并的文档或从多页PDF创建单独的文件。</p>
      <p>您可以按特定页面范围拆分、提取单个页面或按固定间隔分割文档。该工具提供页面的可视预览，使您能够轻松选择所需的内容。</p>
      <p>所有处理都在您的浏览器本地完成，确保您的文档保持私密和安全。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击浏览并选择要拆分的文件。' },
      { step: 2, title: '选择拆分方式', description: '选择拆分方式：按页面范围、提取特定页面或按固定间隔拆分。' },
      { step: 3, title: '定义页面范围', description: '输入要提取的页码或范围（例如：1-5, 8, 10-15）。' },
      { step: 4, title: '拆分并下载', description: '点击拆分创建新的PDF文件，单独下载或作为ZIP压缩包下载。' },
    ],
    useCases: [
      { title: '提取章节', description: '将书籍或手册拆分成单独的章节，便于阅读或分发。', icon: 'book' },
      { title: '分离合并扫描', description: '将批量扫描的文档分成每个原始文档的单独文件。', icon: 'copy' },
      { title: '创建讲义', description: '从演示文稿中提取特定幻灯片或页面以创建重点讲义。', icon: 'presentation' },
    ],
    faq: [
      { question: '可以将PDF拆分成单独的页面吗？', answer: '是的，您可以通过选择"每页拆分"选项将PDF拆分成单独的单页文件。' },
      { question: '拆分时书签会怎样？', answer: '落在提取页面范围内的书签会保留在生成的PDF文件中。' },
      { question: '可以拆分受密码保护的PDF吗？', answer: '您需要先使用我们的解密PDF工具解密PDF，然后再进行拆分。' },
    ],
  },

  'compress-pdf': {
    title: '压缩PDF',
    metaDescription: '减小PDF文件大小同时保持质量。免费在线PDF压缩器，生成更小的文件。',
    keywords: ['压缩pdf', '减小pdf大小', 'pdf压缩器', '缩小pdf', '优化pdf'],
    description: `
      <p>压缩PDF在保持可接受质量的同时减小PDF文档的文件大小。这对于电子邮件附件、网络上传或节省存储空间至关重要。</p>
      <p>该工具提供多种压缩级别，以平衡文件大小减少和质量保持。您可以选择激进压缩以获得最大的大小减少，或选择轻度压缩以保持更高的质量。</p>
      <p>所有压缩都在您的浏览器中进行，确保您的文档永远不会离开您的设备。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要压缩的文档。' },
      { step: 2, title: '选择压缩级别', description: '选择您偏好的压缩级别：低（最佳质量）、中（平衡）或高（最小文件）。' },
      { step: 3, title: '压缩并下载', description: '点击压缩以减小文件大小，然后下载优化后的PDF。' },
    ],
    useCases: [
      { title: '电子邮件附件', description: '减小PDF大小以满足电子邮件附件限制并确保更快的发送。', icon: 'mail' },
      { title: '网络发布', description: '优化PDF以供网络下载，提高页面加载时间和用户体验。', icon: 'globe' },
      { title: '存储优化', description: '压缩存档文档以节省磁盘空间同时保持可访问性。', icon: 'hard-drive' },
    ],
    faq: [
      { question: '可以减少多少文件大小？', answer: '压缩结果因PDF内容而异。图像密集的PDF通常可以减少50-80%，而纯文本PDF可能减少较少。' },
      { question: '压缩会影响文本质量吗？', answer: '文本在所有压缩级别下都保持清晰可读。只有图像和图形会受到压缩影响。' },
      { question: '可以一次压缩多个PDF吗？', answer: '是的，您可以同时上传和压缩最多10个PDF文件。' },
    ],
  },

  'edit-pdf': {
    title: '编辑PDF',
    metaDescription: '在线编辑PDF文件。添加文本、图像、注释、高亮和形状到您的文档。',
    keywords: ['编辑pdf', 'pdf编辑器', '注释pdf', '添加文本到pdf', 'pdf标记'],
    description: `
      <p>编辑PDF提供一套全面的工具来修改和注释您的PDF文档。添加文本、图像、形状、高亮、评论等，无需昂贵的桌面软件。</p>
      <p>直观的编辑器界面使您可以轻松标记文档以供审阅、添加协作注释、编辑敏感信息或用额外内容增强文档。</p>
      <p>所有编辑都在您的浏览器本地进行，确保敏感文档的完全隐私。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要编辑的文档。' },
      { step: 2, title: '选择编辑工具', description: '从工具栏中选择：文本、高亮、形状、图像、评论或编辑工具。' },
      { step: 3, title: '进行编辑', description: '点击文档添加注释，拖动定位元素，使用属性面板进行自定义。' },
      { step: 4, title: '保存并下载', description: '点击保存应用更改并下载编辑后的PDF。' },
    ],
    useCases: [
      { title: '文档审阅', description: '为协作审阅过程添加评论、高亮和标记到文档。', icon: 'message-square' },
      { title: '表单填写', description: '填写文本字段、添加签名并完成PDF表单，无需打印。', icon: 'edit-3' },
      { title: '内容编辑', description: '在共享前永久删除文档中的敏感信息。', icon: 'eye-off' },
    ],
    faq: [
      { question: '可以编辑PDF中的原始文本吗？', answer: '此工具专注于添加注释和新内容。要编辑现有文本，您可能需要使用原始源文档。' },
      { question: '我的编辑是永久的吗？', answer: '注释可以扁平化使其永久，或根据您的偏好保持为可编辑图层。' },
      { question: '可以撤销更改吗？', answer: '是的，编辑器支持撤销/重做功能。您也可以在保存前随时重置为原始文档。' },
    ],
  },

  'jpg-to-pdf': {
    title: 'JPG转PDF',
    metaDescription: '将JPG图像转换为PDF。将多个JPG文件合并成一个PDF文档。',
    keywords: ['jpg转pdf', 'jpeg转pdf', '转换jpg', '图像转pdf', '照片转pdf'],
    description: `
      <p>JPG转PDF可以快速轻松地将您的JPEG图像转换为PDF文档。无论您有单张照片还是多张图像，这个工具都能创建专业外观的PDF文件。</p>
      <p>您可以将多个JPG文件合并成一个PDF，按任意顺序排列，并自定义页面大小和方向。转换过程保留图像质量，同时创建紧凑、可共享的PDF文件。</p>
      <p>所有转换都在您的浏览器中进行，确保您的照片保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传JPG图像', description: '拖放您的JPG文件或点击从设备中选择图像。' },
      { step: 2, title: '排列和配置', description: '通过拖动重新排序图像，选择页面大小和方向选项。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PDF并下载结果。' },
    ],
    useCases: [
      { title: '照片相册', description: '从假期照片或活动照片创建PDF相册，便于分享。', icon: 'image' },
      { title: '文档扫描', description: '将手机拍摄的文档照片转换为正式的PDF文件。', icon: 'camera' },
      { title: '作品集创建', description: '将摄影作品或设计样本编译成专业的PDF作品集。', icon: 'folder' },
    ],
    faq: [
      { question: '可以转换多少张图像？', answer: '您可以将最多100张JPG图像转换成一个PDF文档。' },
      { question: '图像质量会保留吗？', answer: '是的，图像以原始质量嵌入。您可以选择压缩它们以减小文件大小。' },
      { question: '可以为不同图像设置不同的页面大小吗？', answer: '该工具对所有页面应用统一的页面大小。每张图像都会缩放以适应所选页面大小，同时保持纵横比。' },
    ],
  },

  'sign-pdf': {
    title: '签署PDF',
    metaDescription: '为PDF文档添加电子签名。绘制、输入或上传您的签名。',
    keywords: ['签署pdf', '电子签名', '电子签章', 'pdf签名', '数字签名'],
    description: `
      <p>签署PDF允许您快速安全地为PDF文档添加电子签名。通过绘制、输入或上传图像创建您的签名，然后将其放置在文档的任何位置。</p>
      <p>您可以在单个文档中添加多个签名，精确调整大小和位置，并保存签名以供将来使用。该工具非常适合合同、协议、表单和任何需要签名的文档。</p>
      <p>所有签名操作都在您的浏览器本地进行，确保您的文档和签名保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择需要签名的文档。' },
      { step: 2, title: '创建签名', description: '用鼠标或触摸绘制签名，输入姓名生成签名，或上传签名图像。' },
      { step: 3, title: '放置和调整', description: '点击文档放置签名，然后拖动定位和调整大小。' },
      { step: 4, title: '保存并下载', description: '点击保存应用签名并下载已签名的PDF。' },
    ],
    useCases: [
      { title: '合同签署', description: '电子签署合同和协议，无需打印和扫描。', icon: 'file-signature' },
      { title: '表单填写', description: '为申请表、同意书和官方文档添加签名。', icon: 'clipboard' },
      { title: '审批流程', description: '作为审阅和审批流程的一部分签署文档。', icon: 'check-circle' },
    ],
    faq: [
      { question: '电子签名具有法律效力吗？', answer: '电子签名在大多数国家都被法律认可。但是，某些文档可能需要特定类型的数字签名。请查阅当地法规。' },
      { question: '可以保存签名以供将来使用吗？', answer: '是的，您可以将签名保存到浏览器的本地存储中，以便在签署未来文档时快速访问。' },
      { question: '可以在一个文档中添加多个签名吗？', answer: '是的，您可以添加任意数量的签名，在任何页面上独立定位每个签名。' },
    ],
  },

  'crop-pdf': {
    title: '裁剪PDF',
    metaDescription: '裁剪PDF页面以删除边距和不需要的区域。精确修剪PDF文档。',
    keywords: ['裁剪pdf', '修剪pdf', '剪切pdf边距', '调整pdf页面大小', 'pdf裁剪器'],
    description: `
      <p>裁剪PDF允许您修剪边距并从PDF页面中删除不需要的区域。这对于删除多余的空白、聚焦特定内容区域或标准化页面尺寸非常有用。</p>
      <p>您可以统一裁剪所有页面或单独调整每个页面。可视界面准确显示将保留的内容，使您能够轻松获得精确的结果。</p>
      <p>所有裁剪操作都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要裁剪的文档。' },
      { step: 2, title: '定义裁剪区域', description: '拖动裁剪手柄定义要保留的区域，或输入精确的尺寸。' },
      { step: 3, title: '应用到页面', description: '选择将裁剪应用到所有页面或选择特定页面进行裁剪。' },
      { step: 4, title: '裁剪并下载', description: '点击裁剪应用更改并下载裁剪后的PDF。' },
    ],
    useCases: [
      { title: '删除边距', description: '修剪扫描文档或具有大边框的PDF的过多边距。', icon: 'maximize-2' },
      { title: '聚焦内容', description: '裁剪以突出特定内容区域，删除页眉、页脚或侧边栏。', icon: 'target' },
      { title: '标准化页面', description: '通过裁剪到统一尺寸使所有页面大小相同。', icon: 'square' },
    ],
    faq: [
      { question: '裁剪会永久删除内容吗？', answer: '是的，裁剪会删除裁剪区域外的内容。请确保保留原始文件的备份。' },
      { question: '可以对不同页面进行不同的裁剪吗？', answer: '是的，您可以对单个页面或页面组应用不同的裁剪设置。' },
      { question: '裁剪会影响文本质量吗？', answer: '不会，裁剪只删除裁剪边界外的区域。剩余内容保持原始质量。' },
    ],
  },

  'extract-pages': {
    title: '提取页面',
    metaDescription: '从PDF文件中提取特定页面。选择并保存单个页面为新文档。',
    keywords: ['提取pdf页面', '保存pdf页面', '复制pdf页面', 'pdf页面提取器'],
    description: `
      <p>提取页面允许您从PDF文档中选择并保存特定页面为新文件。这非常适合提取相关部分、创建摘录或分离合并的文档。</p>
      <p>您可以提取单个页面、页面范围或多个不连续的页面。可视页面预览使您能够轻松识别和选择所需的页面。</p>
      <p>所有提取操作都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要从中提取页面的文档。' },
      { step: 2, title: '选择页面', description: '点击页面缩略图选择它们，或在输入字段中输入页码和范围。' },
      { step: 3, title: '提取并下载', description: '点击提取创建包含所选页面的新PDF并下载。' },
    ],
    useCases: [
      { title: '创建摘录', description: '从报告或书籍中提取相关页面以创建重点参考文档。', icon: 'file-minus' },
      { title: '分享特定内容', description: '提取特定页面进行分享，无需发送整个文档。', icon: 'share-2' },
      { title: '存档重要页面', description: '提取并保存文档中的关键页面以供长期存档。', icon: 'archive' },
    ],
    faq: [
      { question: '可以提取不连续的页面吗？', answer: '是的，您可以选择任意页面组合，无论是连续的还是分散在整个文档中的。' },
      { question: '书签会保留吗？', answer: '指向提取页面的书签会保留在新文档中。' },
      { question: '可以从多个PDF中提取页面吗？', answer: '此工具一次处理一个PDF。要合并来自多个PDF的页面，请使用合并PDF工具。' },
    ],
  },

  'organize-pdf': {
    title: '整理PDF',
    metaDescription: '重新排序、复制和删除PDF页面。拖放重新整理您的文档。',
    keywords: ['整理pdf', '重新排序pdf页面', '重新排列pdf', 'pdf页面整理器'],
    description: `
      <p>整理PDF提供直观的拖放界面来重新排列PDF文档中的页面。轻松重新排序页面、复制重要部分或删除不需要的页面。</p>
      <p>可视页面缩略图使您能够轻松识别内容并按需排列页面。非常适合重组文档、创建自定义页面顺序或清理扫描文件。</p>
      <p>所有整理操作都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要整理的文档。' },
      { step: 2, title: '重新排列页面', description: '拖动页面缩略图重新排序。根据需要点击每个页面上的复制或删除按钮。' },
      { step: 3, title: '保存并下载', description: '点击保存应用更改并下载重新整理的PDF。' },
    ],
    useCases: [
      { title: '修复页面顺序', description: '纠正扫描或合并错误的页面顺序。', icon: 'arrow-up-down' },
      { title: '创建自定义顺序', description: '为演示文稿或报告按特定顺序排列页面。', icon: 'list' },
      { title: '删除不需要的页面', description: '从文档中删除空白页、重复页或不相关的内容。', icon: 'trash-2' },
    ],
    faq: [
      { question: '可以复制页面吗？', answer: '是的，您可以复制任何页面并将副本放置在文档中的任何位置。' },
      { question: '有撤销功能吗？', answer: '是的，您可以撤销和重做更改。您也可以随时重置为原始顺序。' },
      { question: '可以同时整理多个PDF吗？', answer: '此工具一次处理一个PDF。要合并和整理多个PDF，请先使用合并PDF工具合并它们。' },
    ],
  },

  'delete-pages': {
    title: '删除页面',
    metaDescription: '从PDF文件中删除不需要的页面。轻松选择和删除特定页面。',
    keywords: ['删除pdf页面', '移除pdf页面', 'pdf页面删除器', '从pdf删除页面'],
    description: `
      <p>删除页面允许您快速轻松地从PDF文档中删除不需要的页面。无论您需要删除空白页、过时内容还是敏感信息，这个工具都能简化操作。</p>
      <p>可视页面缩略图帮助您准确识别要删除的页面。您可以删除单个页面或一次删除多个页面。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要从中删除页面的文档。' },
      { step: 2, title: '选择要删除的页面', description: '点击页面缩略图标记要删除的页面，或在输入字段中输入页码。' },
      { step: 3, title: '删除并下载', description: '点击删除移除所选页面并下载更新后的PDF。' },
    ],
    useCases: [
      { title: '删除空白页', description: '通过删除意外包含的空白页来清理文档。', icon: 'file-x' },
      { title: '删除敏感内容', description: '在共享文档前删除包含机密信息的页面。', icon: 'shield' },
      { title: '精简文档', description: '删除过时或不相关的页面以创建更集中的文档。', icon: 'filter' },
    ],
    faq: [
      { question: '可以恢复已删除的页面吗？', answer: '删除在输出文件中是永久的。如果以后可能需要这些页面，请保留原始文档的备份。' },
      { question: '可以一次删除多个页面吗？', answer: '是的，您可以同时选择和删除多个页面。' },
      { question: '删除页面会影响书签吗？', answer: '指向已删除页面的书签将被移除。指向剩余页面的书签会保留。' },
    ],
  },

  // ==================== 编辑与注释 ====================
  'bookmark': {
    title: '编辑书签',
    metaDescription: '添加、编辑和管理PDF书签。为您的文档创建导航结构。',
    keywords: ['pdf书签', '编辑书签', '添加书签', 'pdf导航', '目录'],
    description: `
      <p>编辑书签允许您在PDF文档中创建、修改和整理书签。书签提供快速导航到特定部分的功能，使长文档更易于使用。</p>
      <p>您可以添加新书签、编辑现有书签、重新整理书签层次结构或从外部来源导入书签。这个工具对于创建专业、可导航的文档至关重要。</p>
      <p>所有编辑都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要编辑的文档。' },
      { step: 2, title: '管理书签', description: '添加新书签、编辑现有书签或拖动重新整理层次结构。' },
      { step: 3, title: '保存并下载', description: '点击保存应用更改并下载带有更新书签的PDF。' },
    ],
    useCases: [
      { title: '创建导航', description: '为长文档添加书签，帮助读者快速导航到特定部分。', icon: 'navigation' },
      { title: '整理章节', description: '创建反映文档章节组织的层次书签结构。', icon: 'book-open' },
      { title: '提高可访问性', description: '添加书签使文档更易于访问和用户友好。', icon: 'accessibility' },
    ],
    faq: [
      { question: '可以创建嵌套书签吗？', answer: '是的，您可以创建具有父书签和子书签的层次结构。' },
      { question: '可以从文件导入书签吗？', answer: '是的，您可以从JSON或文本文件导入书签结构。' },
      { question: '书签在所有PDF阅读器中都能工作吗？', answer: '是的，书签是所有主要PDF阅读器都支持的标准PDF功能。' },
    ],
  },

  'table-of-contents': {
    title: '目录',
    metaDescription: '为您的PDF生成目录。从书签创建可点击的导航。',
    keywords: ['pdf目录', '目录生成器', 'pdf索引', '文档导航'],
    description: `
      <p>目录为您的PDF文档生成可导航的目录页面。目录可以从现有书签或自定义条目创建，为读者提供概览和快速导航。</p>
      <p>使用不同的样式、字体和布局自定义外观。生成的目录包含可点击的链接，直接跳转到引用的页面。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '配置目录', description: '选择从书签生成或创建自定义条目。选择样式和位置选项。' },
      { step: 3, title: '生成并下载', description: '点击生成创建目录并下载更新后的PDF。' },
    ],
    useCases: [
      { title: '学术论文', description: '为论文、学位论文和研究报告添加专业目录。', icon: 'graduation-cap' },
      { title: '商业报告', description: '为利益相关者创建具有清晰章节列表的可导航报告。', icon: 'bar-chart' },
      { title: '用户手册', description: '为技术文档和用户指南生成全面的目录。', icon: 'book' },
    ],
    faq: [
      { question: '可以自定义目录外观吗？', answer: '是的，您可以为目录选择不同的样式、字体和布局。' },
      { question: '目录插入在哪里？', answer: '默认情况下，目录插入在文档开头，但您可以选择不同的位置。' },
      { question: '目录条目可以点击吗？', answer: '是的，每个条目都是可点击的链接，导航到相应的页面。' },
    ],
  },

  'page-numbers': {
    title: '页码',
    metaDescription: '为PDF文档添加页码。自定义位置、格式和起始编号。',
    keywords: ['添加页码', 'pdf页码', '编号pdf页面', 'pdf分页'],
    description: `
      <p>页码为您的PDF文档添加可自定义的页码编号。从各种格式、位置和样式中选择以匹配您文档的设计。</p>
      <p>您可以设置起始编号、跳过某些页面并使用不同的编号格式（1, 2, 3 或 i, ii, iii）。非常适合创建具有正确分页的专业文档。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '配置编号', description: '选择位置、格式、起始编号以及要编号的页面。' },
      { step: 3, title: '应用并下载', description: '点击应用添加页码并下载更新后的PDF。' },
    ],
    useCases: [
      { title: '专业文档', description: '为报告、提案和商业文档添加页码。', icon: 'file-text' },
      { title: '学术论文', description: '根据学术格式要求编号页面。', icon: 'graduation-cap' },
      { title: '法律文档', description: '为合同和法律文件添加正确的分页。', icon: 'scale' },
    ],
    faq: [
      { question: '可以跳过第一页吗？', answer: '是的，您可以指定要编号的页面和要跳过的页面，如标题页或封面页。' },
      { question: '有哪些编号格式可用？', answer: '您可以使用阿拉伯数字（1, 2, 3）、罗马数字（i, ii, iii 或 I, II, III）或字母（a, b, c）。' },
      { question: '可以添加"第X页，共Y页"格式吗？', answer: '是的，您可以在编号格式中包含总页数。' },
    ],
  },

  'add-watermark': {
    title: '添加水印',
    metaDescription: '为PDF文件添加文本或图像水印。保护和品牌化您的文档。',
    keywords: ['添加水印', 'pdf水印', '盖章pdf', '品牌pdf', '保护pdf'],
    description: `
      <p>添加水印允许您在PDF文档上放置文本或图像水印。水印可以指示文档状态（草稿、机密）、添加品牌或阻止未经授权的复制。</p>
      <p>自定义水印的位置、大小、不透明度、旋转和颜色。应用到所有页面或选择特定页面。该工具支持文本水印和图像水印。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '创建水印', description: '输入文本或上传图像作为水印。调整位置、大小、不透明度和旋转。' },
      { step: 3, title: '应用并下载', description: '点击应用添加水印并下载更新后的PDF。' },
    ],
    useCases: [
      { title: '文档保护', description: '添加"机密"或"草稿"水印以指示文档状态。', icon: 'shield' },
      { title: '品牌文档', description: '为官方文档添加公司标志或名称。', icon: 'award' },
      { title: '版权声明', description: '添加版权信息以保护知识产权。', icon: 'copyright' },
    ],
    faq: [
      { question: '可以使用图像作为水印吗？', answer: '是的，您可以上传PNG、JPG或SVG图像作为水印。' },
      { question: '可以使水印半透明吗？', answer: '是的，您可以调整不透明度，从完全透明到完全不透明。' },
      { question: '可以对不同页面应用不同的水印吗？', answer: '该工具对选定的页面应用相同的水印。对于不同的水印，需要多次处理文档。' },
    ],
  },

  'header-footer': {
    title: '页眉页脚',
    metaDescription: '为PDF文档添加页眉和页脚。包含页码、日期和自定义文本。',
    keywords: ['pdf页眉', 'pdf页脚', '添加页眉页脚', 'pdf信头'],
    description: `
      <p>页眉页脚为您的PDF文档添加可自定义的页眉和页脚。在页眉或页脚区域包含页码、日期、文档标题或任何自定义文本。</p>
      <p>将内容定位在页眉/页脚的左侧、中间或右侧。如果需要，可以为奇数页和偶数页使用不同的内容。非常适合创建具有一致格式的专业文档。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '配置页眉/页脚', description: '为页眉和页脚区域输入文本。添加页码、日期或自定义文本。' },
      { step: 3, title: '应用并下载', description: '点击应用添加页眉/页脚并下载更新后的PDF。' },
    ],
    useCases: [
      { title: '商业文档', description: '为专业文档添加公司名称和页码。', icon: 'briefcase' },
      { title: '法律文档', description: '在法律文件中包含案件编号、日期和页面引用。', icon: 'scale' },
      { title: '学术论文', description: '添加带有论文标题和作者姓名的页眉。', icon: 'graduation-cap' },
    ],
    faq: [
      { question: '可以在奇数页和偶数页上有不同的页眉吗？', answer: '是的，您可以为奇数页和偶数页配置不同的内容。' },
      { question: '可以包含当前日期吗？', answer: '是的，您可以插入显示当前日期的动态日期字段。' },
      { question: '可以在某些页面上跳过页眉/页脚吗？', answer: '是的，您可以指定哪些页面应该有页眉/页脚，哪些应该跳过。' },
    ],
  },

  'invert-colors': {
    title: '反转颜色',
    metaDescription: '反转PDF颜色以进行暗模式阅读。将文档转换为负片颜色。',
    keywords: ['反转pdf颜色', 'pdf暗模式', '负片pdf', '反转颜色'],
    description: `
      <p>反转颜色可以反转PDF文档中的颜色，创建负片图像效果。这对于创建文档的暗模式版本特别有用，便于在低光条件下阅读。</p>
      <p>该工具可以反转所有颜色或选择性地保留某些元素（如图像）。非常适合在夜间阅读文档时减少眼睛疲劳。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '配置选项', description: '选择是反转所有内容还是保留图像。' },
      { step: 3, title: '反转并下载', description: '点击反转处理文档并下载结果。' },
    ],
    useCases: [
      { title: '夜间阅读', description: '创建文档的暗模式版本，便于夜间舒适阅读。', icon: 'moon' },
      { title: '减少眼睛疲劳', description: '反转明亮的文档以减少长时间阅读时的眼睛疲劳。', icon: 'eye' },
      { title: '节省打印', description: '反转文档以减少打印草稿时的墨水使用。', icon: 'printer' },
    ],
    faq: [
      { question: '图像也会被反转吗？', answer: '默认情况下会。您可以选择在反转文本和背景的同时保留原始图像。' },
      { question: '可以只反转特定页面吗？', answer: '是的，您可以选择要反转的页面。' },
      { question: '反转是可逆的吗？', answer: '您可以再次反转文档以大致恢复原始颜色。' },
    ],
  },

  'background-color': {
    title: '背景颜色',
    metaDescription: '更改PDF背景颜色。为文档页面添加彩色背景。',
    keywords: ['pdf背景颜色', '更改pdf背景', '彩色pdf', 'pdf页面颜色'],
    description: `
      <p>背景颜色允许您更改或添加PDF页面的背景颜色。这可以提高可读性、增加视觉趣味或满足您的品牌要求。</p>
      <p>为背景选择任何颜色并应用到所有页面或选定的页面。该工具在添加背景图层的同时保留所有现有内容。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '选择颜色', description: '使用颜色选择器选择背景颜色或输入十六进制代码。' },
      { step: 3, title: '应用并下载', description: '点击应用添加背景并下载更新后的PDF。' },
    ],
    useCases: [
      { title: '提高可读性', description: '添加浅奶油色或棕褐色背景以减少眼睛疲劳。', icon: 'eye' },
      { title: '品牌文档', description: '使用品牌颜色作为营销材料的背景。', icon: 'palette' },
      { title: '突出部分', description: '使用不同的背景颜色来区分文档部分。', icon: 'layers' },
    ],
    faq: [
      { question: '背景会覆盖现有内容吗？', answer: '不会，背景添加在现有内容后面，保留所有文本和图像。' },
      { question: '可以为不同页面使用不同的颜色吗？', answer: '您需要多次处理文档才能在不同页面上使用不同的颜色。' },
      { question: '可以删除现有背景吗？', answer: '此工具添加背景。要删除背景，您可能需要使用编辑PDF工具。' },
    ],
  },

  'text-color': {
    title: '更改文本颜色',
    metaDescription: '更改PDF文档中的文本颜色。修改所有文本内容的颜色。',
    keywords: ['更改pdf文本颜色', 'pdf文本颜色', '修改文本颜色', '重新着色pdf文本'],
    description: `
      <p>更改文本颜色允许您修改PDF文档中文本的颜色。这对于改善对比度、匹配品牌或创建文档的视觉变体非常有用。</p>
      <p>选择新颜色并应用到文档中的所有文本。该工具处理文本元素的同时保留图像和其他内容。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '选择颜色', description: '使用颜色选择器选择新的文本颜色或输入十六进制代码。' },
      { step: 3, title: '应用并下载', description: '点击应用更改文本颜色并下载更新后的PDF。' },
    ],
    useCases: [
      { title: '改善对比度', description: '更改文本颜色以提高与背景的可读性。', icon: 'contrast' },
      { title: '品牌一致性', description: '更新文本颜色以匹配品牌指南。', icon: 'palette' },
      { title: '无障碍性', description: '调整文本颜色以满足无障碍对比度要求。', icon: 'accessibility' },
    ],
    faq: [
      { question: '所有文本都会被更改吗？', answer: '是的，该工具更改文档中所有文本元素的颜色。' },
      { question: '可以只更改特定文本吗？', answer: '此工具更改所有文本。对于选择性更改，请使用编辑PDF工具。' },
      { question: '格式化的文本（粗体、斜体）会保留吗？', answer: '是的，文本格式会保留；只有颜色会更改。' },
    ],
  },

  'add-stamps': {
    title: '添加印章',
    metaDescription: '为PDF文档添加印章。使用预设或自定义印章进行审批、审阅等。',
    keywords: ['pdf印章', '添加印章', '审批印章', 'pdf橡皮章'],
    description: `
      <p>添加印章允许您在PDF文档上放置印章图像。使用预设印章如"已批准"、"已拒绝"、"草稿"，或上传自定义印章图像。</p>
      <p>将印章定位在页面的任何位置，调整大小，并应用到单个或多个页面。非常适合文档工作流程、审批和状态指示。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '选择印章', description: '选择预设印章或上传自定义印章图像。' },
      { step: 3, title: '定位并应用', description: '点击放置印章，调整位置和大小，然后下载。' },
    ],
    useCases: [
      { title: '文档审批', description: '在审阅工作流程中为文档添加"已批准"或"已拒绝"印章。', icon: 'check-circle' },
      { title: '状态指示', description: '将文档标记为"草稿"、"最终版"或"机密"。', icon: 'tag' },
      { title: '质量控制', description: '添加质检印章以指示检查或审阅完成。', icon: 'clipboard-check' },
    ],
    faq: [
      { question: '有哪些预设印章可用？', answer: '预设包括已批准、已拒绝、草稿、最终版、机密、副本等。' },
      { question: '可以上传自定义印章吗？', answer: '是的，您可以上传PNG或JPG图像作为自定义印章。' },
      { question: '可以在一个文档中添加多个印章吗？', answer: '是的，您可以添加多个印章并独立定位每个印章。' },
    ],
  },

  'remove-annotations': {
    title: '删除注释',
    metaDescription: '从PDF文件中删除注释。删除评论、高亮和标记。',
    keywords: ['删除pdf注释', '删除评论', '删除高亮', '清理pdf'],
    description: `
      <p>删除注释可以从PDF文档中去除评论、高亮、便签和其他注释。这将创建一个没有标记的干净文档版本。</p>
      <p>您可以删除所有注释或选择性地删除特定类型。非常适合创建已审阅文档的最终版本或删除敏感评论。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '选择注释类型', description: '选择要删除的注释类型：评论、高亮、链接等。' },
      { step: 3, title: '删除并下载', description: '点击删除去除注释并下载干净的PDF。' },
    ],
    useCases: [
      { title: '完成文档', description: '在发布最终文档前删除审阅评论和标记。', icon: 'file-check' },
      { title: '隐私保护', description: '在共享前删除可能包含敏感信息的评论。', icon: 'shield' },
      { title: '干净分发', description: '创建已注释文档的干净副本以供分发。', icon: 'copy' },
    ],
    faq: [
      { question: '可以删除哪些类型的注释？', answer: '评论、高亮、下划线、删除线、便签、印章和链接都可以删除。' },
      { question: '可以保留一些注释吗？', answer: '是的，您可以选择要删除的注释类型和要保留的类型。' },
      { question: '这是可逆的吗？', answer: '不，注释删除是永久的。如果需要，请保留原始文件的备份。' },
    ],
  },

  'form-filler': {
    title: '表单填写',
    metaDescription: '在线填写PDF表单。无需打印即可完成交互式PDF表单。',
    keywords: ['填写pdf表单', 'pdf表单填写器', '完成pdf表单', '交互式pdf'],
    description: `
      <p>表单填写允许您直接在浏览器中完成交互式PDF表单。填写文本字段、勾选复选框、选择选项并添加签名，无需打印文档。</p>
      <p>该工具支持标准PDF表单和XFA表单。您填写的数据可以保存，表单可以扁平化以防止进一步编辑。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的表单数据保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF表单', description: '拖放您的PDF表单或点击选择文件。' },
      { step: 2, title: '填写表单', description: '点击表单字段输入文本、勾选复选框或选择选项。' },
      { step: 3, title: '保存并下载', description: '点击保存保留您的输入并下载已填写的表单。' },
    ],
    useCases: [
      { title: '申请表', description: '完成工作申请、许可申请和注册表单。', icon: 'clipboard' },
      { title: '税务表单', description: '电子填写税务文件和财务表单。', icon: 'file-text' },
      { title: '合同', description: '在签署前用您的信息完成合同表单。', icon: 'file-signature' },
    ],
    faq: [
      { question: '可以保存进度吗？', answer: '是的，您可以保存部分填写的表单并稍后继续。' },
      { question: '什么是表单扁平化？', answer: '扁平化将表单字段转换为静态内容，防止进一步编辑。' },
      { question: '支持XFA表单吗？', answer: '是的，该工具支持标准AcroForms和XFA表单。' },
    ],
  },

  'form-creator': {
    title: '表单创建',
    metaDescription: '创建可填写的PDF表单。添加文本字段、复选框和下拉菜单到文档。',
    keywords: ['创建pdf表单', 'pdf表单创建器', '可填写pdf', '添加表单字段'],
    description: `
      <p>表单创建将静态PDF文档转换为交互式可填写表单。添加文本字段、复选框、单选按钮、下拉菜单等以创建专业表单。</p>
      <p>将表单元素拖放到文档上，配置字段属性，创建可以电子填写的表单。非常适合创建申请表、调查问卷和数据收集表单。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择要转换为表单的文档。' },
      { step: 2, title: '添加表单字段', description: '从工具栏选择字段类型并点击将其放置在文档上。' },
      { step: 3, title: '配置并保存', description: '设置字段属性，然后保存并下载可填写的PDF表单。' },
    ],
    useCases: [
      { title: '申请表', description: '创建可填写的工作申请、会员表单和注册表。', icon: 'user-plus' },
      { title: '调查问卷', description: '构建用于数据收集的交互式调查和问卷。', icon: 'clipboard-list' },
      { title: '订单表', description: '创建带有数量字段和复选框的产品订单表。', icon: 'shopping-cart' },
    ],
    faq: [
      { question: '可以添加哪些字段类型？', answer: '文本字段、复选框、单选按钮、下拉菜单、日期选择器和签名字段。' },
      { question: '可以将字段设为必填吗？', answer: '是的，您可以将字段标记为必填并添加验证规则。' },
      { question: '可以添加计算吗？', answer: '可以为数字字段添加基本计算，如求和和平均值。' },
    ],
  },

  'remove-blank-pages': {
    title: '删除空白页',
    metaDescription: '自动检测并从PDF文档中删除空白页。',
    keywords: ['删除空白页', '删除空页', '清理pdf', 'pdf空白页删除器'],
    description: `
      <p>删除空白页自动检测并从PDF文档中删除空页。这对于清理扫描文档、删除分隔页或消除意外包含的空白页非常有用。</p>
      <p>该工具使用智能检测来识别真正的空白页，同时保留内容最少的页面。您可以调整灵敏度阈值来控制什么算作"空白"。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择文档。' },
      { step: 2, title: '调整阈值', description: '如果需要，设置空白检测阈值（默认值适用于大多数文档）。' },
      { step: 3, title: '删除并下载', description: '点击删除删除空白页并下载清理后的PDF。' },
    ],
    useCases: [
      { title: '清理扫描文档', description: '从批量扫描的文档中删除空白页。', icon: 'scan' },
      { title: '删除分隔符', description: '从合并文档中删除空白分隔页。', icon: 'minus' },
      { title: '减小文件大小', description: '删除不必要的空白页以减小文档大小。', icon: 'minimize-2' },
    ],
    faq: [
      { question: '空白检测如何工作？', answer: '该工具分析页面内容，将内容最少或没有可见内容的页面视为空白。' },
      { question: '可以预览哪些页面将被删除吗？', answer: '是的，检测到的空白页在删除前会高亮显示以供审阅。' },
      { question: '如果页面只有页眉/页脚怎么办？', answer: '您可以调整阈值来确定内容最少的页面是否应被视为空白。' },
    ],
  },

  // ==================== 转换为PDF ====================
  'image-to-pdf': {
    title: '图像转PDF',
    metaDescription: '将任何图像转换为PDF。支持JPG、PNG、WebP、BMP、TIFF、SVG和HEIC格式。',
    keywords: ['图像转pdf', '转换图像', '照片转pdf', '图片转pdf'],
    description: `
      <p>图像转PDF将任何格式的图像转换为PDF文档。支持JPG、PNG、WebP、BMP、TIFF、SVG和HEIC格式，使其成为通用的图像转换器。</p>
      <p>将多个图像合并成一个PDF，按任意顺序排列，并自定义页面大小和方向。非常适合创建照片相册、文档存档或作品集。</p>
      <p>所有转换都在您的浏览器中进行，确保您的图像保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传图像', description: '拖放任何支持格式的图像或点击选择文件。' },
      { step: 2, title: '排列和配置', description: '重新排序图像并选择页面大小和方向选项。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PDF并下载结果。' },
    ],
    useCases: [
      { title: '照片集', description: '将来自各种来源的照片合并成一个PDF相册。', icon: 'images' },
      { title: '混合格式文档', description: '将不同格式的图像转换为统一的PDF。', icon: 'file-stack' },
      { title: '存档创建', description: '从图像集创建PDF存档以供长期存储。', icon: 'archive' },
    ],
    faq: [
      { question: '支持哪些图像格式？', answer: '支持JPG、JPEG、PNG、WebP、BMP、TIFF、TIF、SVG、HEIC和HEIF格式。' },
      { question: '可以混合不同的图像格式吗？', answer: '是的，您可以将不同格式的图像合并成一个PDF。' },
      { question: '图像质量会保留吗？', answer: '是的，除非您选择压缩，否则图像以原始质量嵌入。' },
    ],
  },

  'png-to-pdf': {
    title: 'PNG转PDF',
    metaDescription: '将PNG图像转换为PDF。保留透明度并合并多个PNG文件。',
    keywords: ['png转pdf', '转换png', 'png转换器', '透明图像转pdf'],
    description: `
      <p>PNG转PDF将您的PNG图像转换为PDF文档，同时保留透明度。非常适合图形、标志、截图和具有透明背景的图像。</p>
      <p>将多个PNG文件合并成一个PDF，按任意顺序排列，并自定义页面设置。转换过程保持原始图像的高质量。</p>
      <p>所有转换都在您的浏览器中进行，确保您的图像保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PNG文件', description: '拖放您的PNG图像或点击选择文件。' },
      { step: 2, title: '排列和配置', description: '重新排序图像并选择页面大小选项。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PDF并下载。' },
    ],
    useCases: [
      { title: '图形作品集', description: '将PNG图形和设计编译成专业作品集。', icon: 'palette' },
      { title: '截图文档', description: '将截图转换为PDF文档。', icon: 'monitor' },
      { title: '标志集', description: '创建标志和品牌资产的PDF目录。', icon: 'award' },
    ],
    faq: [
      { question: '透明度会保留吗？', answer: 'PNG透明度在PDF输出中会保留。' },
      { question: 'PNG动画怎么办？', answer: '动画PNG使用第一帧转换为静态图像。' },
      { question: '可以设置背景颜色吗？', answer: '是的，您可以为透明区域选择背景颜色。' },
    ],
  },

  'webp-to-pdf': {
    title: 'WebP转PDF',
    metaDescription: '将WebP图像转换为PDF。现代图像格式转换，保持质量。',
    keywords: ['webp转pdf', '转换webp', 'webp转换器', '网络图像转pdf'],
    description: `
      <p>WebP转PDF将现代WebP图像转换为PDF文档。WebP是一种流行的网络图像格式，这个工具使转换这些图像变得简单，便于打印或存档。</p>
      <p>将多个WebP文件合并成一个PDF，并可自定义页面设置。转换过程在创建紧凑PDF文件的同时保留图像质量。</p>
      <p>所有转换都在您的浏览器中进行，确保您的图像保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传WebP文件', description: '拖放您的WebP图像或点击选择文件。' },
      { step: 2, title: '配置选项', description: '排列图像并选择页面大小和方向。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PDF。' },
    ],
    useCases: [
      { title: '网络内容存档', description: '将网络图像转换为PDF以供离线存档。', icon: 'globe' },
      { title: '打印准备', description: '将WebP图像转换为PDF以供打印。', icon: 'printer' },
      { title: '格式标准化', description: '将现代WebP转换为通用兼容的PDF。', icon: 'file-check' },
    ],
    faq: [
      { question: '什么是WebP格式？', answer: 'WebP是Google开发的现代图像格式，为网络图像提供卓越的压缩。' },
      { question: '质量会保留吗？', answer: '是的，转换过程保留原始图像质量。' },
      { question: '可以转换动画WebP吗？', answer: '动画WebP文件转换为静态图像。' },
    ],
  },

  'svg-to-pdf': {
    title: 'SVG转PDF',
    metaDescription: '将SVG矢量图形转换为PDF。保留可缩放性和质量。',
    keywords: ['svg转pdf', '转换svg', '矢量转pdf', '可缩放图形转pdf'],
    description: `
      <p>SVG转PDF将可缩放矢量图形转换为PDF文档，同时保留其矢量质量。SVG文件在任何尺寸下都保持清晰，这种质量在PDF输出中得以保持。</p>
      <p>非常适合转换标志、图标、插图和技术图纸。生成的PDF保持原始矢量图形的可缩放性。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文件保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传SVG文件', description: '拖放您的SVG文件或点击选择。' },
      { step: 2, title: '配置设置', description: '选择页面大小和排列选项。' },
      { step: 3, title: '转换并下载', description: '点击转换创建矢量PDF。' },
    ],
    useCases: [
      { title: '标志转换', description: '将SVG标志转换为PDF以供印刷材料使用。', icon: 'award' },
      { title: '技术图纸', description: '将CAD导出和技术插图转换为PDF。', icon: 'ruler' },
      { title: '图标集', description: '创建图标集和图形的PDF目录。', icon: 'grid' },
    ],
    faq: [
      { question: '矢量质量会保留吗？', answer: '是的，SVG矢量质量在PDF输出中完全保留。' },
      { question: '可以转换复杂的SVG吗？', answer: '是的，支持带有渐变、滤镜和效果的复杂SVG。' },
      { question: '嵌入字体怎么办？', answer: 'SVG文件中的嵌入字体在PDF中会保留。' },
    ],
  },

  'bmp-to-pdf': {
    title: 'BMP转PDF',
    metaDescription: '将BMP位图图像转换为PDF。传统格式支持，保持质量。',
    keywords: ['bmp转pdf', '转换bmp', '位图转pdf', 'bmp转换器'],
    description: `
      <p>BMP转PDF将位图图像转换为PDF文档。BMP是Windows环境中常用的传统图像格式，这个工具使将这些文件转换为现代PDF格式变得简单。</p>
      <p>将多个BMP文件合并成一个PDF，并可自定义设置。转换过程在保持图像质量的同时压缩通常较大的BMP文件。</p>
      <p>所有转换都在您的浏览器中进行，确保您的图像保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传BMP文件', description: '拖放您的BMP图像或点击选择文件。' },
      { step: 2, title: '配置选项', description: '排列图像并选择页面设置。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PDF。' },
    ],
    useCases: [
      { title: '传统文件转换', description: '将旧的BMP文件转换为现代PDF格式。', icon: 'history' },
      { title: 'Windows截图', description: '将Windows位图截图转换为PDF。', icon: 'monitor' },
      { title: '存档现代化', description: '将传统图像存档更新为PDF格式。', icon: 'archive' },
    ],
    faq: [
      { question: '文件大小会减小吗？', answer: '是的，BMP文件在转换为PDF时通常会显著压缩。' },
      { question: '质量会保留吗？', answer: '是的，转换过程中图像质量得以保持。' },
      { question: '支持哪些BMP颜色深度？', answer: '支持所有标准BMP颜色深度，包括24位和32位。' },
    ],
  },

  'heic-to-pdf': {
    title: 'HEIC转PDF',
    metaDescription: '将iPhone HEIC照片转换为PDF。Apple图像格式转换变得简单。',
    keywords: ['heic转pdf', '转换heic', 'iphone照片转pdf', 'apple图像转pdf'],
    description: `
      <p>HEIC转PDF将Apple的高效图像格式照片转换为PDF文档。HEIC是iPhone和iPad上的默认照片格式，这个工具使分享这些照片变得简单。</p>
      <p>将多张HEIC照片合并成一个PDF，非常适合从iPhone照片创建照片相册或文档存档。</p>
      <p>所有转换都在您的浏览器中进行，确保您的照片保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传HEIC文件', description: '拖放您的HEIC照片或点击选择文件。' },
      { step: 2, title: '排列照片', description: '重新排序照片并选择页面设置。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PDF。' },
    ],
    useCases: [
      { title: 'iPhone照片相册', description: '从iPhone照片创建PDF相册以供分享。', icon: 'smartphone' },
      { title: '文档扫描', description: '将iPhone文档扫描转换为PDF格式。', icon: 'scan' },
      { title: '跨平台分享', description: '将HEIC转换为PDF以实现通用兼容性。', icon: 'share-2' },
    ],
    faq: [
      { question: '什么是HEIC格式？', answer: 'HEIC（高效图像容器）是Apple的图像格式，比JPEG提供更好的压缩。' },
      { question: '支持实况照片吗？', answer: '实况照片使用关键帧转换为静态图像。' },
      { question: 'EXIF数据会保留吗？', answer: '照片元数据可以在转换过程中选择性保留或删除。' },
    ],
  },

  'tiff-to-pdf': {
    title: 'TIFF转PDF',
    metaDescription: '将TIFF图像转换为PDF。支持多页TIFF文件和高质量转换。',
    keywords: ['tiff转pdf', '转换tiff', 'tif转pdf', '多页tiff'],
    description: `
      <p>TIFF转PDF将TIFF图像（包括多页TIFF文件）转换为PDF文档。TIFF常用于高质量扫描和专业图形。</p>
      <p>多页TIFF文件自动转换为多页PDF。转换过程保留原始图像的高质量。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文件保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传TIFF文件', description: '拖放您的TIFF文件或点击选择。' },
      { step: 2, title: '配置选项', description: '选择页面设置和压缩选项。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PDF。' },
    ],
    useCases: [
      { title: '扫描文档', description: '将高质量扫描从TIFF转换为PDF。', icon: 'scan' },
      { title: '专业图形', description: '转换专业TIFF图形以供分发。', icon: 'image' },
      { title: '存档转换', description: '将TIFF存档转换为更易访问的PDF格式。', icon: 'archive' },
    ],
    faq: [
      { question: '支持多页TIFF吗？', answer: '是的，多页TIFF文件自动转换为多页PDF。' },
      { question: '质量会保留吗？', answer: '是的，TIFF质量在PDF输出中完全保留。' },
      { question: '使用什么压缩？', answer: '您可以选择无损和有损压缩选项。' },
    ],
  },

  'txt-to-pdf': {
    title: '文本转PDF',
    metaDescription: '将纯文本文件转换为PDF。自定义字体、边距和页面布局。',
    keywords: ['txt转pdf', '文本转pdf', '转换文本文件', '纯文本转pdf'],
    description: `
      <p>文本转PDF将纯文本文件转换为格式化的PDF文档。自定义字体、大小、边距和页面布局，从简单文本创建专业外观的文档。</p>
      <p>非常适合转换代码文件、日志、笔记或任何纯文本内容为可共享的PDF格式。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文件保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传文本文件', description: '拖放您的.txt文件或点击选择。' },
      { step: 2, title: '自定义格式', description: '选择字体、大小、边距和页面设置。' },
      { step: 3, title: '转换并下载', description: '点击转换创建格式化的PDF。' },
    ],
    useCases: [
      { title: '代码文档', description: '将源代码文件转换为PDF以供文档使用。', icon: 'code' },
      { title: '日志存档', description: '将日志文件转换为PDF以供存档。', icon: 'file-text' },
      { title: '笔记转换', description: '将纯文本笔记转换为格式化的PDF文档。', icon: 'sticky-note' },
    ],
    faq: [
      { question: '有哪些字体可用？', answer: '有多种字体可用，包括用于代码的等宽字体。' },
      { question: '自动换行吗？', answer: '是的，长行会自动换行以适应页面。' },
      { question: '可以保留格式吗？', answer: '原始文本中的空白和缩进会保留。' },
    ],
  },

  'json-to-pdf': {
    title: 'JSON转PDF',
    metaDescription: '将JSON文件转换为格式化的PDF。语法高亮和结构化输出。',
    keywords: ['json转pdf', '转换json', 'json查看器', 'json格式化器'],
    description: `
      <p>JSON转PDF将JSON数据文件转换为格式化、可读的PDF文档。输出包括语法高亮和正确的缩进，便于阅读。</p>
      <p>非常适合记录API响应、配置文件或任何需要以可读格式共享或存档的JSON数据。</p>
      <p>所有转换都在您的浏览器中进行，确保您的数据保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传JSON文件', description: '拖放您的.json文件或点击选择。' },
      { step: 2, title: '配置显示', description: '选择格式化选项和语法高亮。' },
      { step: 3, title: '转换并下载', description: '点击转换创建格式化的PDF。' },
    ],
    useCases: [
      { title: 'API文档', description: '将API响应转换为PDF以供文档使用。', icon: 'code' },
      { title: '配置存档', description: '以可读的PDF格式存档配置文件。', icon: 'settings' },
      { title: '数据报告', description: '从JSON数据导出创建PDF报告。', icon: 'bar-chart' },
    ],
    faq: [
      { question: '包含语法高亮吗？', answer: '是的，JSON语法用颜色高亮显示键、值和类型。' },
      { question: '嵌套数据如何处理？', answer: '嵌套对象和数组会正确缩进以提高可读性。' },
      { question: '大型JSON文件怎么办？', answer: '大文件会自动分页到多个页面。' },
    ],
  },

  // ==================== 从PDF转换 ====================
  'pdf-to-jpg': {
    title: 'PDF转JPG',
    metaDescription: '将PDF页面转换为JPG图像。高质量提取，可自定义分辨率。',
    keywords: ['pdf转jpg', 'pdf转jpeg', '转换pdf为图像', '提取pdf图像'],
    description: `
      <p>PDF转JPG将PDF文档页面转换为高质量的JPG图像。提取所有页面或选择特定页面进行转换，可自定义分辨率和质量设置。</p>
      <p>非常适合从PDF中提取图像、创建缩略图或转换文档以供网络使用。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择页面和质量', description: '选择要转换的页面并设置质量/DPI选项。' },
      { step: 3, title: '转换并下载', description: '点击转换提取图像并作为ZIP下载。' },
    ],
    useCases: [
      { title: '网络发布', description: '将PDF页面转换为图像以供网站使用。', icon: 'globe' },
      { title: '社交媒体', description: '提取页面作为图像以供社交媒体分享。', icon: 'share-2' },
      { title: '演示文稿', description: '将PDF幻灯片转换为图像以供演示。', icon: 'presentation' },
    ],
    faq: [
      { question: '有哪些质量设置可用？', answer: '您可以设置72到300的DPI和1-100的JPEG质量。' },
      { question: '可以只转换特定页面吗？', answer: '是的，您可以选择单个页面或页面范围进行转换。' },
      { question: '多个页面如何处理？', answer: '每个页面成为单独的JPG文件，作为ZIP压缩包下载。' },
    ],
  },

  'pdf-to-png': {
    title: 'PDF转PNG',
    metaDescription: '将PDF页面转换为PNG图像。无损质量，支持透明度。',
    keywords: ['pdf转png', '转换pdf为png', 'pdf图像提取', '无损pdf转换'],
    description: `
      <p>PDF转PNG将PDF文档页面转换为具有无损压缩的高质量PNG图像。PNG格式完美保留图像质量并支持透明度。</p>
      <p>非常适合提取图形、图表或任何质量保持至关重要的内容。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '配置选项', description: '选择页面并设置分辨率（DPI）选项。' },
      { step: 3, title: '转换并下载', description: '点击转换提取PNG图像。' },
    ],
    useCases: [
      { title: '图形提取', description: '以完美质量提取图表和图形。', icon: 'image' },
      { title: '设计资产', description: '将PDF设计转换为PNG以供编辑软件使用。', icon: 'palette' },
      { title: '文档', description: '为技术文档创建高质量图像。', icon: 'file-text' },
    ],
    faq: [
      { question: '为什么选择PNG而不是JPG？', answer: 'PNG提供无损压缩和透明度支持，非常适合图形和文本。' },
      { question: '支持透明背景吗？', answer: '是的，具有透明度的PDF页面在PNG输出中会保留。' },
      { question: '应该使用什么DPI？', answer: '屏幕查看使用150 DPI，打印使用300 DPI。' },
    ],
  },

  'pdf-to-webp': {
    title: 'PDF转WebP',
    metaDescription: '将PDF页面转换为WebP图像。现代格式，出色的压缩。',
    keywords: ['pdf转webp', '转换pdf为webp', '现代图像格式', '网络优化图像'],
    description: `
      <p>PDF转WebP将PDF文档页面转换为WebP图像，这是Google的现代图像格式，提供出色的压缩和高质量。</p>
      <p>WebP图像比JPG或PNG更小，同时保持相当的质量，非常适合网络使用。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '设置质量选项', description: '选择页面并设置质量/压缩设置。' },
      { step: 3, title: '转换并下载', description: '点击转换创建WebP图像。' },
    ],
    useCases: [
      { title: '网络优化', description: '从PDF内容创建网络优化的图像。', icon: 'globe' },
      { title: '带宽节省', description: '减小图像文件大小以加快加载速度。', icon: 'zap' },
      { title: '现代网站', description: '为现代网络项目使用现代图像格式。', icon: 'layout' },
    ],
    faq: [
      { question: '什么是WebP格式？', answer: 'WebP是Google的现代图像格式，提供卓越的压缩。' },
      { question: 'WebP被广泛支持吗？', answer: '是的，所有现代浏览器都支持WebP格式。' },
      { question: 'WebP文件小多少？', answer: 'WebP文件通常比同等JPG文件小25-35%。' },
    ],
  },

  'pdf-to-bmp': {
    title: 'PDF转BMP',
    metaDescription: '将PDF页面转换为BMP位图图像。未压缩格式，最大兼容性。',
    keywords: ['pdf转bmp', '转换pdf为位图', '未压缩图像', '传统格式'],
    description: `
      <p>PDF转BMP将PDF文档页面转换为BMP位图图像。BMP是一种未压缩格式，确保与传统系统和应用程序的最大兼容性。</p>
      <p>虽然BMP文件比压缩格式大，但它们提供完美的质量和通用兼容性。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择页面', description: '选择要转换的页面并设置DPI。' },
      { step: 3, title: '转换并下载', description: '点击转换创建BMP图像。' },
    ],
    useCases: [
      { title: '传统系统', description: '创建与旧软件兼容的图像。', icon: 'history' },
      { title: 'Windows应用程序', description: '为Windows特定应用程序生成BMP文件。', icon: 'monitor' },
      { title: '未压缩存档', description: '从PDF创建未压缩的图像存档。', icon: 'archive' },
    ],
    faq: [
      { question: '为什么使用BMP格式？', answer: 'BMP提供未压缩的质量和与传统系统的最大兼容性。' },
      { question: 'BMP文件更大吗？', answer: '是的，BMP文件未压缩，比JPG或PNG大得多。' },
      { question: '支持哪些颜色深度？', answer: '支持24位和32位颜色深度。' },
    ],
  },

  'pdf-to-tiff': {
    title: 'PDF转TIFF',
    metaDescription: '将PDF转换为TIFF图像。专业质量，支持多页。',
    keywords: ['pdf转tiff', '转换pdf为tiff', '专业图像', '多页tiff'],
    description: `
      <p>PDF转TIFF将PDF文档转换为高质量的TIFF图像。由于其无损压缩，TIFF是专业打印和存档的首选格式。</p>
      <p>创建单页TIFF或将所有页面合并成多页TIFF文件。非常适合专业和存档目的。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '配置输出', description: '选择单页或多页TIFF并设置DPI。' },
      { step: 3, title: '转换并下载', description: '点击转换创建TIFF图像。' },
    ],
    useCases: [
      { title: '专业打印', description: '从PDF文档创建可打印的TIFF文件。', icon: 'printer' },
      { title: '文档存档', description: '以高质量TIFF格式存档文档。', icon: 'archive' },
      { title: '出版', description: '将PDF转换为TIFF以供出版工作流程使用。', icon: 'book' },
    ],
    faq: [
      { question: '可以创建多页TIFF吗？', answer: '是的，您可以将所有PDF页面合并成一个多页TIFF。' },
      { question: '有哪些压缩选项？', answer: 'LZW、ZIP和无压缩选项可用。' },
      { question: '打印应该使用什么DPI？', answer: '专业打印使用300 DPI或更高。' },
    ],
  },

  'pdf-to-greyscale': {
    title: 'PDF转灰度',
    metaDescription: '将彩色PDF转换为灰度。减小文件大小并准备黑白打印。',
    keywords: ['pdf转灰度', '灰度pdf', '黑白pdf', '删除颜色'],
    description: `
      <p>PDF转灰度将彩色PDF文档转换为灰度（黑白）。这可以减小文件大小并为黑白打印准备文档。</p>
      <p>转换过程在删除颜色信息的同时保留文本清晰度和图像细节。非常适合草稿打印或创建打印机友好版本。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的彩色PDF文件或点击选择。' },
      { step: 2, title: '预览转换', description: '预览灰度版本的外观。' },
      { step: 3, title: '转换并下载', description: '点击转换创建灰度PDF。' },
    ],
    useCases: [
      { title: '节省打印', description: '转换为灰度以节省彩色打印成本。', icon: 'printer' },
      { title: '草稿文档', description: '创建黑白草稿以供审阅。', icon: 'file-text' },
      { title: '减小文件大小', description: '通过删除颜色信息减小PDF大小。', icon: 'minimize-2' },
    ],
    faq: [
      { question: '文本会保持可读吗？', answer: '是的，灰度转换过程中文本清晰度得以保留。' },
      { question: '文件会小多少？', answer: '文件大小减少因情况而异，但对于颜色密集的文档可以减少20-50%。' },
      { question: '可以只转换特定页面吗？', answer: '是的，您可以选择要转换为灰度的页面。' },
    ],
  },

  'pdf-to-json': {
    title: 'PDF转JSON',
    metaDescription: '将PDF内容提取为JSON格式。从PDF文档获取结构化数据。',
    keywords: ['pdf转json', '提取pdf数据', 'pdf解析器', '结构化pdf数据'],
    description: `
      <p>PDF转JSON将PDF文档中的内容提取为结构化的JSON格式。提取文本、元数据、页面信息和文档结构以供程序化使用。</p>
      <p>非常适合数据提取、文档分析或将PDF内容集成到应用程序和工作流程中。</p>
      <p>所有提取都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择要提取的数据', description: '选择要提取的内容：文本、元数据、结构。' },
      { step: 3, title: '提取并下载', description: '点击提取生成JSON并下载。' },
    ],
    useCases: [
      { title: '数据提取', description: '从PDF文档中提取结构化数据。', icon: 'database' },
      { title: '文档分析', description: '以编程方式分析PDF结构和内容。', icon: 'search' },
      { title: '集成', description: '通过JSON将PDF内容导入应用程序。', icon: 'plug' },
    ],
    faq: [
      { question: '提取哪些数据？', answer: '文本内容、元数据、页面尺寸、字体和文档结构。' },
      { question: 'JSON格式有文档吗？', answer: '是的，JSON模式是一致且有良好文档的。' },
      { question: '可以从扫描的PDF中提取吗？', answer: '扫描的PDF需要先进行OCR。在提取前使用我们的OCR PDF工具。' },
    ],
  },

  'pdf-to-pptx': {
    title: 'PDF转PowerPoint',
    metaDescription: '将PDF转换为PowerPoint (PPTX)演示文稿。每一页转为高质量幻灯片。',
    keywords: ['pdf转ppt', 'pdf转pptx', 'pdf转powerpoint', 'pdf演示文稿'],
    description: `
      <p>PDF转PowerPoint将您的PDF文档转换为可编辑的PowerPoint演示文稿(PPTX)。每个PDF页面都会转换为保持完美视觉布局的高质量幻灯片。</p>
      <p>此工具非常适合将报告、讲义或任何PDF内容转换为演示格式。</p>
      <p>所有转换都在您的浏览器中进行，确保您的文档隐私安全。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择质量', description: '选择幻灯片的图像质量(DPI)。' },
      { step: 3, title: '转换并下载', description: '点击转换创建PowerPoint演示文稿。' },
    ],
    useCases: [
      { title: '创建演示', description: '将PDF文档转换为会议幻灯片。', icon: 'presentation' },
      { title: '培训材料', description: '将教材转换为交互式演示文稿。', icon: 'book-open' },
      { title: '内容复用', description: '将现有内容转换为可编辑幻灯片。', icon: 'refresh-cw' },
    ],
    faq: [
      { question: '幻灯片可编辑吗？', answer: '每张幻灯片包含PDF页面的图像。您可以在顶部添加内容。' },
      { question: '应该选什么DPI？', answer: '屏幕展示选150 DPI，打印选300 DPI。' },
      { question: '支持多页吗？', answer: '是的，每一页都会成为一张单独的幻灯片。' },
    ],
  },

  'pdf-to-excel': {
    title: 'PDF转Excel',
    metaDescription: '将PDF转换为Excel表格。将表格提取为XLSX格式。',
    keywords: ['pdf转excel', 'pdf转xlsx', '提取表格', 'pdf数据提取'],
    description: `
      <p>PDF转Excel将您的PDF文档转换为可编辑的Excel电子表格(XLSX)。工具自动检测并提取表格。</p>
      <p>非常适合分析财务报告或数据表。每页的表格提取到单独的Sheet中。</p>
      <p>所有转换都在您的浏览器中进行，确保您的数据隐私安全。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '处理', description: '工具将自动识别表格。' },
      { step: 3, title: '下载Excel', description: '下载包含提取表格的文件。' },
    ],
    useCases: [
      { title: '财务分析', description: '转换银行账单或发票。', icon: 'trending-up' },
      { title: '数据提取', description: '从报告中提取数据表。', icon: 'database' },
      { title: '清单转换', description: '将PDF库存清单转换为表格。', icon: 'clipboard' },
    ],
    faq: [
      { question: '如何处理表格？', answer: '每页的表格提取到对应的Excel工作表中。' },
      { question: '如果没有表格？', answer: '将创建一个提示信息工作表。' },
      { question: '保留格式吗？', answer: '数据保留，视觉格式可能简化。' },
    ],
  },

  // ==================== 整理与管理 ====================
  'ocr-pdf': {
    title: 'OCR PDF',
    metaDescription: '使用OCR使扫描的PDF可搜索。从图像和扫描文档中提取文本。',
    keywords: ['ocr pdf', '可搜索pdf', '文本识别', '扫描转文本'],
    description: `
      <p>OCR PDF使用光学字符识别从PDF中的扫描文档和图像中提取文本。将基于图像的PDF转换为可搜索、可选择文本的文档。</p>
      <p>支持多种语言，确保无论文档语言如何都能准确识别文本。在添加可搜索文本层的同时保留原始布局。</p>
      <p>所有OCR处理都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传扫描的PDF', description: '拖放您的扫描PDF或点击选择。' },
      { step: 2, title: '选择语言', description: '选择文档语言以获得准确识别。' },
      { step: 3, title: '处理并下载', description: '点击处理运行OCR并下载可搜索的PDF。' },
    ],
    useCases: [
      { title: '数字化存档', description: '使扫描的文档存档可搜索。', icon: 'archive' },
      { title: '文档搜索', description: '在扫描文档中启用文本搜索。', icon: 'search' },
      { title: '文本提取', description: '从扫描文档中提取文本以供编辑。', icon: 'type' },
    ],
    faq: [
      { question: '支持哪些语言？', answer: '支持100多种语言，包括英语、中文、日语、韩语等。' },
      { question: '原始布局会保留吗？', answer: '是的，原始视觉布局会保留，并添加可搜索的文本层。' },
      { question: 'OCR有多准确？', answer: '准确性取决于扫描质量，但对于清晰的文档通常超过95%。' },
    ],
  },

  'alternate-merge': {
    title: '交替合并',
    metaDescription: '通过交替页面合并PDF。将正面和背面扫描合并成一个文档。',
    keywords: ['交替合并', '交错pdf', '合并扫描', '正反面合并'],
    description: `
      <p>交替合并通过交替交错两个PDF的页面来合并它们。这非常适合将分别扫描的正面和背面页面合并成一个文档。</p>
      <p>上传两个PDF，工具将通过交替从每个PDF中取一页来合并它们。您还可以反转其中一个文档的顺序以适应从后到前的扫描。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传两个PDF', description: '上传正面页面PDF和背面页面PDF。' },
      { step: 2, title: '配置顺序', description: '选择是否为从后到前的扫描反转第二个文档。' },
      { step: 3, title: '合并并下载', description: '点击合并交错页面并下载。' },
    ],
    useCases: [
      { title: '双面扫描', description: '合并分别扫描的正面和背面页面。', icon: 'copy' },
      { title: '文档组装', description: '交错来自两个相关文档的页面。', icon: 'layers' },
      { title: '书籍扫描', description: '将奇数页和偶数页扫描合并成完整的书籍。', icon: 'book' },
    ],
    faq: [
      { question: '如果文档页数不同怎么办？', answer: '较长文档的额外页面会附加在末尾。' },
      { question: '可以反转页面顺序吗？', answer: '是的，您可以在合并前反转任一文档。' },
      { question: '这与普通合并有什么不同？', answer: '是的，普通合并是追加文档；交替合并是交错页面。' },
    ],
  },

  'add-attachments': {
    title: '添加附件',
    metaDescription: '在PDF文档中嵌入文件。将任何文件类型附加到您的PDF。',
    keywords: ['pdf附件', '嵌入文件', '附加到pdf', 'pdf组合'],
    description: `
      <p>添加附件将任何类型的文件嵌入到您的PDF文档中。附加电子表格、图像、源文件或任何其他文档以创建全面的PDF包。</p>
      <p>附件嵌入在PDF中，收件人可以使用任何PDF阅读器提取。非常适合将相关文件一起分发。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文件保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '添加附件', description: '选择要附加到PDF的文件。' },
      { step: 3, title: '保存并下载', description: '点击保存嵌入附件并下载。' },
    ],
    useCases: [
      { title: '项目包', description: '将项目文件与文档PDF捆绑在一起。', icon: 'package' },
      { title: '报告分发', description: '将源数据文件附加到报告PDF。', icon: 'paperclip' },
      { title: '合同包', description: '在合同中包含支持文档。', icon: 'file-text' },
    ],
    faq: [
      { question: '可以附加哪些文件类型？', answer: '任何文件类型都可以附加到PDF。' },
      { question: '有大小限制吗？', answer: '包括附件在内的PDF总大小不应超过500MB。' },
      { question: '收件人可以提取附件吗？', answer: '是的，任何PDF阅读器都可以提取嵌入的附件。' },
    ],
  },

  'extract-attachments': {
    title: '提取附件',
    metaDescription: '从PDF中提取嵌入的文件。从PDF文档下载所有附件。',
    keywords: ['提取附件', 'pdf附件', '下载嵌入文件', 'pdf提取'],
    description: `
      <p>提取附件从PDF文档中检索所有嵌入的文件。单独下载附件或作为包含所有文件的ZIP压缩包下载。</p>
      <p>非常适合访问嵌入在PDF包中的源文件、数据或补充材料。</p>
      <p>所有提取都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '查看附件', description: '查看PDF中所有嵌入文件的列表。' },
      { step: 3, title: '提取并下载', description: '下载单个文件或全部作为ZIP。' },
    ],
    useCases: [
      { title: '访问源文件', description: '从PDF报告中提取原始数据文件。', icon: 'download' },
      { title: '恢复附件', description: '从PDF包中检索嵌入的文件。', icon: 'folder-open' },
      { title: '批量提取', description: '一次从多个PDF中提取附件。', icon: 'layers' },
    ],
    faq: [
      { question: '如果没有附件怎么办？', answer: '如果没有找到嵌入文件，工具会提示。' },
      { question: '支持所有附件类型吗？', answer: '是的，所有嵌入的文件类型都可以提取。' },
      { question: '可以从多个PDF中提取吗？', answer: '是的，您可以处理多个PDF并下载所有附件。' },
    ],
  },

  'edit-attachments': {
    title: '编辑附件',
    metaDescription: '管理PDF附件。查看、重命名和删除嵌入的文件。',
    keywords: ['编辑附件', '管理pdf文件', '删除附件', '重命名附件'],
    description: `
      <p>编辑附件让您管理PDF文档中的嵌入文件。查看所有附件、重命名它们或从PDF中删除不需要的文件。</p>
      <p>非常适合在分发前清理PDF包或更新附件信息。</p>
      <p>所有编辑都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '管理附件', description: '查看、重命名或删除嵌入的文件。' },
      { step: 3, title: '保存并下载', description: '点击保存应用更改并下载。' },
    ],
    useCases: [
      { title: '清理PDF', description: '从PDF包中删除不必要的附件。', icon: 'trash-2' },
      { title: '重命名文件', description: '更新附件名称以提高清晰度。', icon: 'edit' },
      { title: '审阅内容', description: '在分发前审核嵌入的文件。', icon: 'eye' },
    ],
    faq: [
      { question: '可以在这里添加新附件吗？', answer: '使用添加附件工具嵌入新文件。' },
      { question: '删除是永久的吗？', answer: '是的，删除的附件无法从输出文件中恢复。' },
      { question: '可以预览附件吗？', answer: '您可以看到文件名和大小；使用提取附件查看内容。' },
    ],
  },

  'divide-pages': {
    title: '分割页面',
    metaDescription: '将PDF页面分割成多个部分。水平或垂直分割页面。',
    keywords: ['分割pdf页面', '拆分页面', '剪切pdf页面', '页面部分'],
    description: `
      <p>分割页面将单个PDF页面分割成多个部分。水平、垂直或网格分割页面，从一页创建多页。</p>
      <p>非常适合分割每页包含多个项目的扫描文档，或将大幅面页面分割成标准尺寸。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '设置分割', description: '选择水平、垂直或网格分割并设置部分数量。' },
      { step: 3, title: '分割并下载', description: '点击分割拆分页面并下载。' },
    ],
    useCases: [
      { title: '分割扫描', description: '分割包含多个文档的扫描页面。', icon: 'scissors' },
      { title: '调整页面大小', description: '将大页面分割成标准纸张尺寸。', icon: 'maximize-2' },
      { title: '创建卡片', description: '将页面分割成卡片大小的部分以供打印。', icon: 'grid' },
    ],
    faq: [
      { question: '可以分割成不等的部分吗？', answer: '目前分割是等分的。使用裁剪PDF进行自定义部分。' },
      { question: '分割线处的内容会怎样？', answer: '内容在分割线处被分割；确保重要内容不在边界处。' },
      { question: '可以只分割特定页面吗？', answer: '是的，您可以选择要分割的页面。' },
    ],
  },

  'add-blank-page': {
    title: '添加空白页',
    metaDescription: '在PDF文档中插入空白页。在任何位置添加空页。',
    keywords: ['添加空白页', '插入页面', '空页', 'pdf页面插入'],
    description: `
      <p>添加空白页在PDF文档的任何位置插入空页。在现有页面之前、之后或之间添加页面，可自定义页面大小。</p>
      <p>非常适合添加笔记空间、创建章节分隔符或为打印准备文档。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择位置', description: '选择在哪里插入空白页以及插入多少页。' },
      { step: 3, title: '添加并下载', description: '点击添加插入页面并下载。' },
    ],
    useCases: [
      { title: '笔记空间', description: '添加空白页以供手写笔记。', icon: 'edit-3' },
      { title: '章节分隔符', description: '在文档章节之间插入空白页。', icon: 'minus' },
      { title: '打印准备', description: '添加页面以对齐双面打印。', icon: 'printer' },
    ],
    faq: [
      { question: '可以选择页面大小吗？', answer: '是的，空白页可以匹配现有页面或使用自定义尺寸。' },
      { question: '可以添加多个空白页吗？', answer: '是的，您可以一次添加任意数量的空白页。' },
      { question: '可以添加彩色页面吗？', answer: '添加空白页后使用背景颜色工具添加颜色。' },
    ],
  },

  'reverse-pages': {
    title: '反转页面',
    metaDescription: '反转PDF页面顺序。将文档页面从最后翻到最前。',
    keywords: ['反转pdf', '翻转页面顺序', '倒置页面', '反转文档'],
    description: `
      <p>反转页面翻转PDF文档中页面的顺序，将最后一页放在最前，第一页放在最后。对于以相反顺序扫描的文档或特定打印需求非常有用。</p>
      <p>该工具处理整个文档或选定的页面范围，保持所有内容和格式。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择页面', description: '选择反转所有页面或特定范围。' },
      { step: 3, title: '反转并下载', description: '点击反转翻转页面顺序并下载。' },
    ],
    useCases: [
      { title: '修复扫描顺序', description: '纠正以相反顺序扫描的文档。', icon: 'refresh-cw' },
      { title: '打印准备', description: '为特定打印要求反转页面。', icon: 'printer' },
      { title: '文档重新排序', description: '快速翻转文档顺序以供审阅。', icon: 'arrow-up-down' },
    ],
    faq: [
      { question: '书签会更新吗？', answer: '是的，书签会更新以指向正确的反转页面。' },
      { question: '可以只反转部分页面吗？', answer: '是的，您可以选择要反转的页面范围。' },
      { question: '这和旋转一样吗？', answer: '不，反转改变页面顺序；旋转改变页面方向。' },
    ],
  },

  'rotate-pdf': {
    title: '旋转PDF',
    metaDescription: '旋转PDF页面。将页面旋转90、180或270度。',
    keywords: ['旋转pdf', '转动pdf页面', 'pdf旋转', '修复方向'],
    description: `
      <p>旋转PDF将文档中的页面旋转90、180或270度。修复方向不正确的扫描、旋转横向页面或调整页面方向以供查看。</p>
      <p>统一旋转所有页面或选择特定页面单独旋转。该工具保留所有内容和格式。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择旋转', description: '选择旋转角度和要旋转的页面。' },
      { step: 3, title: '旋转并下载', description: '点击旋转应用更改并下载。' },
    ],
    useCases: [
      { title: '修复扫描', description: '纠正扫描文档的方向。', icon: 'rotate-cw' },
      { title: '横向页面', description: '旋转横向页面以正确查看。', icon: 'monitor' },
      { title: '混合方向', description: '标准化混合文档中的页面方向。', icon: 'layout' },
    ],
    faq: [
      { question: '可以对不同页面进行不同的旋转吗？', answer: '是的，您可以对不同页面应用不同的旋转。' },
      { question: '旋转会影响打印质量吗？', answer: '不会，旋转保留所有内容质量。' },
      { question: '可以按自定义角度旋转吗？', answer: '旋转限于90度增量（90、180、270）。' },
    ],
  },

  'n-up-pdf': {
    title: 'N合一PDF',
    metaDescription: '每张纸打印多个PDF页面。创建2合1、4合1或自定义布局。',
    keywords: ['n合一pdf', '每张多页', '2合1打印', '页面拼版'],
    description: `
      <p>N合一PDF将多个页面排列到单张纸上，创建2合1、4合1、6合1、9合1或自定义布局。非常适合打印时节省纸张或创建讲义。</p>
      <p>从预设布局中选择或创建自定义排列。该工具自动缩放和定位页面以获得最佳效果。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择布局', description: '选择2合1、4合1、6合1、9合1或自定义网格。' },
      { step: 3, title: '创建并下载', description: '点击创建生成N合一PDF并下载。' },
    ],
    useCases: [
      { title: '节省纸张', description: '每张纸打印多页以减少纸张使用。', icon: 'leaf' },
      { title: '创建讲义', description: '从演示幻灯片制作紧凑的讲义。', icon: 'file-text' },
      { title: '审阅文档', description: '以缩小尺寸打印文档以供审阅。', icon: 'eye' },
    ],
    faq: [
      { question: '有哪些布局可用？', answer: '2合1、4合1、6合1、9合1和自定义网格布局可用。' },
      { question: '可以在页面之间添加边框吗？', answer: '是的，您可以在页面之间添加边框和间距。' },
      { question: '页面顺序会保留吗？', answer: '是的，页面按阅读顺序排列（从左到右，从上到下）。' },
    ],
  },

  'combine-single-page': {
    title: '合并为单页',
    metaDescription: '将PDF页面拼接成一个连续页面。创建可滚动的单页文档。',
    keywords: ['合并页面', '单页pdf', '拼接页面', '连续滚动'],
    description: `
      <p>合并为单页将所有PDF页面拼接成一个连续页面。创建非常适合网络查看或连续阅读的可滚动文档。</p>
      <p>页面垂直连接，间距可自定义。结果是包含所有内容的单个长页面。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '设置间距', description: '选择拼接页面之间的间隙。' },
      { step: 3, title: '合并并下载', description: '点击合并创建单页PDF。' },
    ],
    useCases: [
      { title: '网络文档', description: '创建可滚动的PDF以供网页嵌入。', icon: 'globe' },
      { title: '连续阅读', description: '将分页文档转换为连续滚动。', icon: 'scroll' },
      { title: '长篇内容', description: '合并页面以实现无缝的长篇阅读。', icon: 'file-text' },
    ],
    faq: [
      { question: '有页数限制吗？', answer: '非常长的文档可能受浏览器内存限制。' },
      { question: '可以在页面之间添加分隔符吗？', answer: '是的，您可以在原始页面之间添加间距或线条。' },
      { question: '这适合打印吗？', answer: '结果最适合屏幕查看；打印布局请使用N合一。' },
    ],
  },

  'view-metadata': {
    title: '查看元数据',
    metaDescription: '查看PDF文档属性。查看作者、标题、日期和其他元数据。',
    keywords: ['pdf元数据', '文档属性', 'pdf信息', '查看pdf详情'],
    description: `
      <p>查看元数据显示PDF文件中的所有文档属性和元数据。查看作者、标题、主题、关键词、创建日期、修改日期等。</p>
      <p>对于审核文档、检查文件信息或验证文档真实性非常有用。</p>
      <p>所有查看都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '查看属性', description: '查看以有组织格式显示的所有元数据。' },
      { step: 3, title: '如需导出', description: '可选择将元数据导出为JSON。' },
    ],
    useCases: [
      { title: '文档审核', description: '审阅文档属性以确保合规性。', icon: 'clipboard-check' },
      { title: '验证真实性', description: '检查创建日期和作者信息。', icon: 'shield' },
      { title: '文件信息', description: '获取PDF文件的详细信息。', icon: 'info' },
    ],
    faq: [
      { question: '显示哪些元数据？', answer: '标题、作者、主题、关键词、创建者、生产者、日期和PDF版本。' },
      { question: '可以在这里编辑元数据吗？', answer: '使用编辑元数据工具修改文档属性。' },
      { question: '包含XMP元数据吗？', answer: '是的，标准和XMP元数据都会显示。' },
    ],
  },

  'edit-metadata': {
    title: '编辑元数据',
    metaDescription: '编辑PDF文档属性。更改标题、作者、主题和关键词。',
    keywords: ['编辑pdf元数据', '更改pdf属性', 'pdf作者', '文档信息'],
    description: `
      <p>编辑元数据允许您修改PDF文件中的文档属性。更改标题、作者、主题、关键词和其他元数据字段。</p>
      <p>非常适合纠正文档信息、添加正确的归属或为分发准备文件。</p>
      <p>所有编辑都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '编辑属性', description: '修改标题、作者、主题、关键词和其他字段。' },
      { step: 3, title: '保存并下载', description: '点击保存应用更改并下载。' },
    ],
    useCases: [
      { title: '添加归属', description: '设置正确的作者和创建者信息。', icon: 'user' },
      { title: 'SEO优化', description: '添加关键词和描述以提高可搜索性。', icon: 'search' },
      { title: '文档准备', description: '在共享前准备具有正确元数据的文档。', icon: 'file-check' },
    ],
    faq: [
      { question: '可以编辑哪些字段？', answer: '标题、作者、主题、关键词、创建者和生产者字段。' },
      { question: '可以清除所有元数据吗？', answer: '使用删除元数据工具去除所有文档属性。' },
      { question: '日期可以编辑吗？', answer: '创建和修改日期会自动更新。' },
    ],
  },

  'pdf-to-zip': {
    title: 'PDF转ZIP',
    metaDescription: '将多个PDF打包成ZIP压缩包。压缩和捆绑PDF文件。',
    keywords: ['pdf转zip', '压缩pdf', '捆绑pdf', '存档pdf'],
    description: `
      <p>PDF转ZIP将多个PDF文件打包成一个ZIP压缩包。压缩和捆绑您的PDF以便于共享、存储或备份。</p>
      <p>该工具创建包含所有PDF文件的压缩存档，减少总大小并简化文件管理。</p>
      <p>所有处理都在您的浏览器中进行，确保您的文件保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF', description: '拖放多个PDF文件或点击选择。' },
      { step: 2, title: '配置存档', description: '可选设置存档名称和压缩级别。' },
      { step: 3, title: '创建并下载', description: '点击创建生成ZIP压缩包。' },
    ],
    useCases: [
      { title: '文件共享', description: '捆绑多个PDF以便于共享。', icon: 'share-2' },
      { title: '创建备份', description: '创建PDF集合的压缩备份。', icon: 'archive' },
      { title: '电子邮件附件', description: '将PDF合并成一个附件以供电子邮件使用。', icon: 'mail' },
    ],
    faq: [
      { question: '应用多少压缩？', answer: 'ZIP压缩通常将总大小减少10-30%。' },
      { question: '有文件限制吗？', answer: '您可以在单个存档中包含最多100个PDF。' },
      { question: '可以设置密码吗？', answer: '目前不支持创建受密码保护的ZIP。' },
    ],
  },

  'compare-pdfs': {
    title: '比较PDF',
    metaDescription: '比较两个PDF文档。高亮显示版本之间的差异。',
    keywords: ['比较pdf', 'pdf差异', '文档比较', '版本比较'],
    description: `
      <p>比较PDF分析两个PDF文档并高亮显示它们之间的差异。非常适合审阅文档修订、检查合同更改或验证编辑。</p>
      <p>以并排或叠加模式查看文档，差异会高亮显示。该工具识别文本更改、添加和删除。</p>
      <p>所有比较都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传两个PDF', description: '上传原始和修改后的PDF文档。' },
      { step: 2, title: '比较文档', description: '以并排或叠加模式查看高亮显示的差异。' },
      { step: 3, title: '导出结果', description: '下载比较报告或带注释的PDF。' },
    ],
    useCases: [
      { title: '合同审阅', description: '比较合同版本以识别更改。', icon: 'file-text' },
      { title: '文档修订', description: '审阅文档版本之间的编辑。', icon: 'git-compare' },
      { title: '质量保证', description: '验证只进行了预期的更改。', icon: 'check-circle' },
    ],
    faq: [
      { question: '检测哪些类型的差异？', answer: '文本添加、删除、修改和格式更改。' },
      { question: '可以比较扫描的文档吗？', answer: '扫描的文档应先进行OCR处理以进行文本比较。' },
      { question: '有视觉比较吗？', answer: '是的，叠加模式显示页面之间的视觉差异。' },
    ],
  },

  'posterize-pdf': {
    title: '海报化PDF',
    metaDescription: '将大型PDF页面分割成可打印的瓷砖。从PDF页面创建海报。',
    keywords: ['海报化pdf', '瓷砖pdf', '大幅面打印', 'pdf海报'],
    description: `
      <p>海报化PDF将大型PDF页面分割成可以在标准纸张上打印并组装成海报的较小瓷砖。非常适合打印大型图表、地图或艺术品。</p>
      <p>配置网格大小和重叠以便于组装。该工具自动计算目标输出尺寸的瓷砖尺寸。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的大幅面PDF或点击选择。' },
      { step: 2, title: '配置瓷砖', description: '设置网格大小、重叠和输出纸张尺寸。' },
      { step: 3, title: '创建并下载', description: '点击创建生成可打印的瓷砖。' },
    ],
    useCases: [
      { title: '海报打印', description: '在标准纸张上打印大型海报。', icon: 'maximize-2' },
      { title: '地图打印', description: '分段打印大型地图以供组装。', icon: 'map' },
      { title: '艺术品复制', description: '从PDF艺术品创建大型打印品。', icon: 'image' },
    ],
    faq: [
      { question: '应该使用多少重叠？', answer: '建议10-20mm的重叠以便于组装时对齐。' },
      { question: '可以添加裁切标记吗？', answer: '是的，可以添加裁切标记以帮助切割和对齐。' },
      { question: '支持哪些纸张尺寸？', answer: '支持A4、Letter、A3和自定义尺寸。' },
    ],
  },

  // ==================== 优化与修复 ====================
  'fix-page-size': {
    title: '修复页面大小',
    metaDescription: '标准化PDF页面大小。将所有页面转换为统一尺寸。',
    keywords: ['修复页面大小', '标准化pdf', '统一页面', '调整pdf页面大小'],
    description: `
      <p>修复页面大小将PDF中的所有页面标准化为统一尺寸。将混合尺寸文档转换为一致的页面大小，以便专业展示或打印。</p>
      <p>从标准尺寸（A4、Letter等）中选择或设置自定义尺寸。内容会缩放或定位以适应新的页面大小。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择目标大小', description: '选择标准尺寸或输入自定义尺寸。' },
      { step: 3, title: '应用并下载', description: '点击应用标准化页面并下载。' },
    ],
    useCases: [
      { title: '打印准备', description: '标准化页面以实现一致的打印。', icon: 'printer' },
      { title: '文档清理', description: '修复页面大小不一致的文档。', icon: 'file-check' },
      { title: '专业文档', description: '创建统一的文档以供分发。', icon: 'briefcase' },
    ],
    faq: [
      { question: '内容如何处理？', answer: '内容会缩放以适应或居中在新页面大小上。' },
      { question: '可以保持纵横比吗？', answer: '是的，内容可以按比例缩放以适应。' },
      { question: '有哪些标准尺寸可用？', answer: 'A4、A3、Letter、Legal和其他常见尺寸。' },
    ],
  },

  'linearize-pdf': {
    title: '线性化PDF',
    metaDescription: '优化PDF以实现快速网络查看。启用渐进式加载。',
    keywords: ['线性化pdf', '快速网络查看', '优化pdf', '渐进式pdf'],
    description: `
      <p>线性化PDF优化您的文档以实现快速网络查看。线性化的PDF可以在整个文件下载完成之前开始显示，改善用户体验。</p>
      <p>也称为"快速网络查看"，此优化重新组织PDF结构以实现网络浏览器中的渐进式加载。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '线性化', description: '点击线性化优化以供网络查看。' },
      { step: 3, title: '下载', description: '下载优化后的PDF。' },
    ],
    useCases: [
      { title: '网络发布', description: '优化PDF以供网站下载。', icon: 'globe' },
      { title: '电子邮件附件', description: '创建为收件人更快打开的PDF。', icon: 'mail' },
      { title: '在线文档', description: '改善在线文档的查看体验。', icon: 'cloud' },
    ],
    faq: [
      { question: '什么是线性化？', answer: '线性化重新组织PDF数据以实现渐进式加载。' },
      { question: '会减小文件大小吗？', answer: '线性化可能由于添加的结构而略微增加文件大小。' },
      { question: '与所有查看器兼容吗？', answer: '是的，线性化的PDF在所有PDF阅读器中都能工作。' },
    ],
  },

  'page-dimensions': {
    title: '页面尺寸',
    metaDescription: '分析PDF页面大小。查看文档中所有页面的尺寸。',
    keywords: ['pdf页面大小', '页面尺寸', 'pdf测量', '文档大小'],
    description: `
      <p>页面尺寸分析并显示PDF文档中每个页面的大小。以各种单位（英寸、毫米、点）查看尺寸并识别非标准大小的页面。</p>
      <p>对于打印准备、文档分析或识别不一致的页面大小非常有用。</p>
      <p>所有分析都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '查看尺寸', description: '查看所有页面显示的页面大小。' },
      { step: 3, title: '导出报告', description: '可选择将尺寸导出为JSON。' },
    ],
    useCases: [
      { title: '打印规划', description: '打印前检查页面大小。', icon: 'printer' },
      { title: '文档分析', description: '识别尺寸异常的页面。', icon: 'search' },
      { title: '质量控制', description: '验证页面大小符合规格。', icon: 'check-circle' },
    ],
    faq: [
      { question: '有哪些单位可用？', answer: '英寸、毫米、厘米和点。' },
      { question: '会显示方向吗？', answer: '是的，会指示纵向或横向方向。' },
      { question: '可以修复不一致的大小吗？', answer: '使用修复页面大小工具标准化尺寸。' },
    ],
  },

  'remove-restrictions': {
    title: '删除限制',
    metaDescription: '删除PDF限制。解锁打印、复制和编辑权限。',
    keywords: ['删除pdf限制', '解锁pdf', 'pdf权限', '取消pdf限制'],
    description: `
      <p>删除限制解锁具有权限限制的PDF，这些限制阻止打印、复制或编辑。此工具在保留文档内容的同时删除所有者密码限制。</p>
      <p>注意：此工具无法删除阻止打开文档的用户密码。对于受密码保护的文件，请使用解密PDF。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传受限PDF', description: '拖放您的受限PDF或点击选择。' },
      { step: 2, title: '删除限制', description: '点击删除解锁文档。' },
      { step: 3, title: '下载', description: '下载不受限制的PDF。' },
    ],
    useCases: [
      { title: '启用打印', description: '解锁阻止打印的PDF。', icon: 'printer' },
      { title: '启用复制', description: '允许文本选择和复制。', icon: 'copy' },
      { title: '启用编辑', description: '删除文档编辑限制。', icon: 'edit' },
    ],
    faq: [
      { question: '这合法吗？', answer: '从您拥有或有权使用的文档中删除限制通常是合法的。' },
      { question: '可以删除打开密码吗？', answer: '不能，对于受密码保护的文档请使用解密PDF。' },
      { question: '内容会受影响吗？', answer: '不会，只删除限制；内容保持不变。' },
    ],
  },

  'repair-pdf': {
    title: '修复PDF',
    metaDescription: '修复损坏的PDF文件。恢复和修复受损文档。',
    keywords: ['修复pdf', '修复pdf', '恢复pdf', '损坏的pdf'],
    description: `
      <p>修复PDF尝试修复损坏或受损的PDF文件。该工具分析文档结构并重建它以尽可能多地恢复内容。</p>
      <p>对于恢复无法打开、显示错误或由于损坏而缺少内容的文件非常有用。</p>
      <p>所有修复都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传损坏的PDF', description: '拖放您损坏的PDF或点击选择。' },
      { step: 2, title: '修复文档', description: '点击修复尝试恢复。' },
      { step: 3, title: '下载', description: '如果成功，下载修复后的PDF。' },
    ],
    useCases: [
      { title: '恢复文件', description: '恢复无法正常打开的PDF。', icon: 'refresh-cw' },
      { title: '修复错误', description: '修复显示错误消息的文件。', icon: 'wrench' },
      { title: '恢复内容', description: '从部分损坏的文件中恢复内容。', icon: 'file-check' },
    ],
    faq: [
      { question: '所有PDF都可以修复吗？', answer: '成功取决于损坏的类型和程度。' },
      { question: '所有内容都会恢复吗？', answer: '该工具尽可能多地恢复；严重损坏的文件可能有损失。' },
      { question: '应该保留原件吗？', answer: '是的，始终保留原始文件作为备份。' },
    ],
  },

  // ==================== 安全PDF ====================
  'encrypt-pdf': {
    title: '加密PDF',
    metaDescription: '为PDF文件添加密码保护。添加加密并设置权限。',
    keywords: ['加密pdf', '密码保护pdf', '安全pdf', 'pdf加密'],
    description: `
      <p>加密PDF为您的PDF文档添加密码保护和加密。设置用户密码以防止打开，设置所有者密码以控制打印和复制等权限。</p>
      <p>根据不同的安全需求选择不同的加密级别（128位或256位AES）。</p>
      <p>所有加密都在您的浏览器中进行，确保您的密码和文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '设置密码', description: '输入用户密码和/或所有者密码。配置权限。' },
      { step: 3, title: '加密并下载', description: '点击加密保护您的PDF并下载。' },
    ],
    useCases: [
      { title: '机密文档', description: '保护敏感的商业文档。', icon: 'lock' },
      { title: '个人文件', description: '保护个人文档如税务申报表。', icon: 'shield' },
      { title: '受控分发', description: '限制收件人对文档的操作。', icon: 'key' },
    ],
    faq: [
      { question: '用户密码和所有者密码有什么区别？', answer: '用户密码防止打开；所有者密码控制权限。' },
      { question: '使用什么加密？', answer: '提供128位或256位AES加密选项。' },
      { question: '可以只设置权限而不设置用户密码吗？', answer: '是的，您可以只设置所有者密码来控制权限。' },
    ],
  },

  'sanitize-pdf': {
    title: '清理PDF',
    metaDescription: '从PDF中删除隐藏数据。清除元数据、脚本和敏感信息。',
    keywords: ['清理pdf', '清洁pdf', '删除隐藏数据', 'pdf隐私'],
    description: `
      <p>清理PDF从您的文档中删除隐藏数据和潜在敏感信息。去除元数据、嵌入脚本、附件、评论和其他隐藏内容。</p>
      <p>对于准备公开分发的文档或当隐私是关注点时至关重要。</p>
      <p>所有清理都在您的浏览器中进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '选择要删除的内容', description: '选择要去除的隐藏数据类型。' },
      { step: 3, title: '清理并下载', description: '点击清理清洁PDF并下载。' },
    ],
    useCases: [
      { title: '公开发布', description: '准备文档以供公开分发。', icon: 'globe' },
      { title: '隐私保护', description: '在共享前删除个人信息。', icon: 'shield' },
      { title: '安全合规', description: '满足文档处理的安全要求。', icon: 'check-circle' },
    ],
    faq: [
      { question: '删除哪些隐藏数据？', answer: '元数据、脚本、附件、评论、表单数据和隐藏图层。' },
      { question: '可见内容会受影响吗？', answer: '不会，只删除隐藏数据；可见内容保持不变。' },
      { question: '这是可逆的吗？', answer: '不，删除的数据无法恢复。保留原件的备份。' },
    ],
  },

  'decrypt-pdf': {
    title: '解密PDF',
    metaDescription: '从PDF文件中删除密码。解锁受密码保护的文档。',
    keywords: ['解密pdf', '删除pdf密码', '解锁pdf', 'pdf密码删除器'],
    description: `
      <p>解密PDF从PDF文档中删除密码保护。输入当前密码以解锁文件并创建不受保护的副本。</p>
      <p>此工具要求您知道当前密码。它无法破解或绕过未知密码。</p>
      <p>所有解密都在您的浏览器中进行，确保您的密码和文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传受保护的PDF', description: '拖放您受密码保护的PDF。' },
      { step: 2, title: '输入密码', description: '输入当前文档密码。' },
      { step: 3, title: '解密并下载', description: '点击解密删除保护并下载。' },
    ],
    useCases: [
      { title: '删除旧密码', description: '当不再需要密码时解锁文档。', icon: 'unlock' },
      { title: '简化访问', description: '创建不受保护的副本以便于共享。', icon: 'share-2' },
      { title: '存档文档', description: '在长期存档前删除密码。', icon: 'archive' },
    ],
    faq: [
      { question: '可以破解未知密码吗？', answer: '不能，您必须知道当前密码才能解密。' },
      { question: '原始文件会被修改吗？', answer: '不会，会创建一个新的不受保护的副本。' },
      { question: '如果忘记密码怎么办？', answer: '很遗憾，我们无法恢复忘记的密码。' },
    ],
  },

  'flatten-pdf': {
    title: '扁平化PDF',
    metaDescription: '扁平化PDF表单和注释。使内容不可编辑。',
    keywords: ['扁平化pdf', '扁平化表单', '扁平化注释', '不可编辑pdf'],
    description: `
      <p>扁平化PDF将表单字段和注释等交互元素转换为静态内容。扁平化的PDF看起来相同，但不能再编辑。</p>
      <p>非常适合完成已填写的表单、保留注释或创建不可编辑的文档版本。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您带有表单或注释的PDF。' },
      { step: 2, title: '选择要扁平化的内容', description: '选择扁平化表单、注释或两者。' },
      { step: 3, title: '扁平化并下载', description: '点击扁平化创建静态PDF。' },
    ],
    useCases: [
      { title: '完成表单', description: '锁定已填写的表单数据以防止更改。', icon: 'lock' },
      { title: '保留注释', description: '使注释在文档中永久化。', icon: 'check-circle' },
      { title: '存档文档', description: '创建不可编辑的版本以供存档。', icon: 'archive' },
    ],
    faq: [
      { question: '扁平化是可逆的吗？', answer: '不，扁平化是永久的。保留原件的备份。' },
      { question: '外观会改变吗？', answer: '不会，文档看起来相同，但不再是交互式的。' },
      { question: '会减小文件大小吗？', answer: '有时会，因为交互元素被转换为更简单的内容。' },
    ],
  },

  'remove-metadata': {
    title: '删除元数据',
    metaDescription: '从PDF文件中去除元数据。删除作者、日期和文档属性。',
    keywords: ['删除pdf元数据', '去除元数据', 'pdf隐私', '匿名pdf'],
    description: `
      <p>删除元数据从您的PDF文件中去除所有文档属性和元数据。删除作者姓名、创建日期、软件信息和其他识别数据。</p>
      <p>在共享文档或当元数据可能泄露敏感信息时，对于隐私至关重要。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '删除元数据', description: '点击删除去除所有元数据。' },
      { step: 3, title: '下载', description: '下载无元数据的PDF。' },
    ],
    useCases: [
      { title: '隐私保护', description: '在共享前删除个人信息。', icon: 'shield' },
      { title: '匿名文档', description: '创建没有作者归属的文档。', icon: 'user-x' },
      { title: '干净分发', description: '分发没有内部元数据的文档。', icon: 'send' },
    ],
    faq: [
      { question: '删除哪些元数据？', answer: '作者、标题、主题、关键词、日期、创建者和生产者信息。' },
      { question: 'XMP元数据会删除吗？', answer: '是的，标准和XMP元数据都会去除。' },
      { question: '内容会受影响吗？', answer: '不会，只删除元数据；文档内容保持不变。' },
    ],
  },

  'change-permissions': {
    title: '更改权限',
    metaDescription: '修改PDF权限。控制打印、复制和编辑访问。',
    keywords: ['pdf权限', '更改pdf访问', '限制pdf', 'pdf安全'],
    description: `
      <p>更改权限修改PDF文档的访问控制。启用或禁用打印、复制、编辑和注释权限。</p>
      <p>设置所有者密码以强制执行这些限制。收件人可以查看文档，但在可执行的操作上受到限制。</p>
      <p>所有处理都在您的浏览器本地进行，确保您的文档保持私密。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放您的PDF文件或点击选择。' },
      { step: 2, title: '设置权限', description: '启用或禁用打印、复制、编辑和注释。' },
      { step: 3, title: '应用并下载', description: '设置所有者密码并下载受限PDF。' },
    ],
    useCases: [
      { title: '防止复制', description: '禁用文本复制以保护内容。', icon: 'copy' },
      { title: '控制打印', description: '限制或允许文档打印。', icon: 'printer' },
      { title: '限制编辑', description: '防止对文档的修改。', icon: 'edit-3' },
    ],
    faq: [
      { question: '需要密码吗？', answer: '需要所有者密码来强制执行权限。' },
      { question: '权限可以删除吗？', answer: '是的，使用所有者密码或删除限制工具。' },
      { question: '所有PDF阅读器都兼容吗？', answer: '大多数PDF阅读器尊重权限，但有些可能不强制执行。' },
    ],
  }, 'pdf-to-docx': {
    title: 'PDF转Word',
    metaDescription: '将PDF转换为可编辑的Word文档（DOCX）。保留原始布局、格式和图像。',
    keywords: ['pdf转word', 'pdf转docx', 'pdf转可编辑文档', 'pdf转换器'],
    description: `
      <p>PDF转Word工具可将您的PDF文档转换为完全可编辑的Microsoft Word (DOCX)文件。该工具采用先进的解析技术，能够最大限度地保留原始文档的排版、字体、表格和图像。</p>
      <p>无需重新打字，即可轻松修改PDF内容。非常适合处理合同、报告、简历以及任何需要深度编辑的文档。</p>
      <p>所有转换均在您的浏览器本地完成，确保您的商业机密和个人隐私不会被泄露。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放PDF文件或点击选择要转换的文档。' },
      { step: 2, title: '开始转换', description: '等待系统自动解析并重建文档结构。' },
      { step: 3, title: '下载Word文档', description: '点击下载生成的DOCX文件，并在Microsoft Word或WPS中打开。' },
    ],
    useCases: [
      { title: '合同修改', description: '将PDF格式的合同转回Word，以便进行条款修订和版本对比。', icon: 'file-text' },
      { title: '简历更新', description: '找回以前制作的PDF简历，转换为Word格式快速更新工作经历。', icon: 'user' },
      { title: '资料整理', description: '从大型PDF报告中提取文本和表格，用于撰写新的文档或分析报告。', icon: 'copy' },
    ],
    faq: [
      { question: '转换后的排版会乱吗？', answer: '对于标准文档，我们的算法能实现极高的还原度。但如果原PDF是由图片生成的扫描件，建议先使用OCR工具。' },
      { question: '支持WPS或Google Docs吗？', answer: '生成的.docx文件是国际标准格式，完全兼容Microsoft Word、WPS Office、Google Docs和LibreOffice。' },
      { question: '转换受保护的PDF吗？', answer: '如果PDF设置了打开权限，您需要先使用"解密PDF"工具移除密码。' },
    ],
  },



  'pdf-to-txt': {
    title: 'PDF转文本',
    metaDescription: '从PDF中提取纯文本。移除所有格式和图像，获取最简洁的文字内容。',
    keywords: ['pdf转txt', 'pdf提取文本', '获取pdf文字', 'pdf纯文本'],
    description: `
      <p>PDF转文本工具旨在为您提供最纯粹的文字提取体验。它会剥离文档中的背景、图像、链接和复杂的排版，仅保留最核心的文字内容。</p>
      <p>适合需要将PDF内容导入文本编辑器、进行代码分析或准备机器翻译语料的用户。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF文件', description: '拖放PDF文件到此处。' },
      { step: 2, title: '提取文字', description: '系统将快速扫描所有页面的字符流。' },
      { step: 3, title: '下载文本文件', description: '获取.txt格式的纯文本文件。' },
    ],
    useCases: [
      { title: '电子书转文本', description: '将PDF格式的小说转换为纯文本，方便在旧款电子书阅读器上使用。', icon: 'book' },
      { title: '语料库建设', description: '批量从PDF文档中提取文字内容，用于AI训练或大数据分析。', icon: 'code' },
      { title: '翻译准备', description: '提取纯文本内容，方便导入CAT工具或网页翻译器。', icon: 'languages' },
    ],
    faq: [
      { question: '扫描件能转成文本吗？', answer: '普通转换工具无法处理扫描件，请点击导航栏中的"OCR PDF"进行识别。' },
      { question: '排版会乱吗？', answer: 'TXT不支持样式，但我们会尽力通过空格和换行保留原始文本的逻辑顺序。' },
      { question: '支持特殊字符吗？', answer: '支持。提取出的文本默认采用UTF-8编码，兼容中文、韩文、日文等全球语言。' },
    ],
  },

  'pdf-to-pdfa': {
    title: 'PDF转PDF/A',
    metaDescription: '将普通PDF转换为适合长期存档的PDF/A格式。符合ISO标准。',
    keywords: ['pdf转pdfa', 'pdf长期存档', '符合性转换', 'iso标准pdf'],
    description: `
      <p>PDF/A是PDF的ISO标准化版本，专门用于电子文档的长期保存和存档。它确保了文档在未来几十年内即使软件环境发生变化，其显示效果依然保持一致。</p>
      <p>该工具会嵌入所有字体并移除不符合存档规范的动态元素（如JavaScript），使文档变得更加稳健和透明。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF', description: '选择需要永久存档的重要文件。' },
      { step: 2, title: '转换标准', description: '系统将自动调整文档属性以符合PDF/A-1b、2b或3b规范。' },
      { step: 3, title: '下载存档文件', description: '获取适用于政府、法律或学术要求的存档PDF。' },
    ],
    useCases: [
      { title: '法律文书存档', description: '将合同和裁决书转换为PDF/A，确保长期司法效力。', icon: 'scale' },
      { title: '政府公文递交', description: '满足政府部门对递交文件必须为PDF/A格式的要求。', icon: 'landmark' },
      { title: '论文提交', description: '高校图书馆通常要求毕业论文采用PDF/A格式以进行永久馆藏。', icon: 'graduation-cap' },
    ],
    faq: [
      { question: 'PDF/A有什么好处？', answer: '它具有自包含性，这意味着显示文档所需的所有信息（如字体）都已保存在文件内，不依赖外部链接。' },
      { question: '普通PDF查看器能打开吗？', answer: '完全可以。PDF/A与所有现有的PDF阅读器完美兼容。' },
      { question: '转换后文件会变大吗？', answer: '通常会，因为必须嵌入所有字体文件以确保长期显示的准确性。' },
    ],
  },

  'pdf-to-html': {
    title: 'PDF转HTML',
    metaDescription: '将PDF页面转换为网页格式（HTML）。支持自适应布局和跨平台浏览。',
    keywords: ['pdf转html', 'pdf转网页', 'pdf在线发布', 'pdf发布为网页'],
    description: `
      <p>PDF转HTML工具可以将您的静态PDF文档转化为可直接在浏览器中浏览的网页。转换后的内容支持文字检索，并能自适应不同的屏幕尺寸。</p>
      <p>非常适合将PDF手册、宣传册或研究论文发布到网站上，提供比下载PDF文件更好的用户体验。</p>
    `,
    howToUse: [
      { step: 1, title: '上传PDF', description: '选择要网页化的PDF文件。' },
      { step: 2, title: '生成HTML', description: '系统将页面结构、样式和图片重新编码为HTML/CSS。' },
      { step: 3, title: '下载或查看', description: '下载包含HTML文件和资源的压缩包。' },
    ],
    useCases: [
      { title: '在线展示手册', description: '将产品手册转换为网页，方便客户直接通过手机浏览器查阅。', icon: 'monitor' },
      { title: 'SEO增强', description: '将PDF内容转为HTML页面，更容易被搜索引擎索引，提高网站流量。', icon: 'search' },
      { title: '内容分发', description: '制作可以在不安装PDF阅读器的情况下就能查看的轻量化内容。', icon: 'share-2' },
    ],
    faq: [
      { question: '转换后的网页支持响应式吗？', answer: '是的，我们生成的代码能适应手机、平板和桌面设备。' },
      { question: '图片能保留吗？', answer: '可以，PDF中的所有插图和照片都会被优化并保存为网页适用的图像格式。' },
      { question: 'HTML文件里会有乱码吗？', answer: '不会。系统会正确映射编码，确保转换后的文字内容准确无误。' },
    ],
  },
};
