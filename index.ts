import {
  Packer,
  Document,
  Paragraph,
  TextRun,
  ExternalHyperlink,
  SectionType,
} from 'docx'
import * as fs from 'fs'
import type { fill_in_strings } from './interface_fill_in_strings.ts'
import { defaults } from './secrets.ts'
import { input } from '@inquirer/prompts'

function get_underscore_company_name(company_name: string) {
  let underscore_name = ''
  company_name.split('').forEach((character) => {
    if (character === ' ') {
      underscore_name += '_'
    } else {
      underscore_name += character
    }
  })
  return underscore_name
}

async function set_options(defaults: fill_in_strings) {
  defaults.company_name = await input({
    message: 'What is the company name?',
    default: defaults.company_name,
  })
  defaults.role = await input({
    message: 'What is the role you are applying for?',
    default: defaults.role,
  })
  defaults.company_projects = await input({
    message: "What is the company's project?",
    default: defaults.company_projects,
  })
  const current_date = new Date()
  defaults.date = await input({
    message: "What is today's date?",
    default: current_date.toDateString(),
  })
  return defaults
}

async function make_document(fill_ins: fill_in_strings) {
  const company_name = fill_ins.company_name
  const email = fill_ins.email
  const address = fill_ins.address
  const phone = fill_ins.phone
  const date = fill_ins.date
  const role = fill_ins.role
  const company_projects = fill_ins.company_projects
  const output_directory = fill_ins.output_directory
  const underscore_company_name = get_underscore_company_name(company_name)
  const document_title = `Atchison_Skyler_${underscore_company_name}_Cover_Letter`
  const email_link = new ExternalHyperlink({
    children: [
      new TextRun({
        text: `${email}`,
        style: 'Hyperlink',
      }),
    ],
    link: `mailto:${email}`,
  })

  const doc = new Document({
    creator: 'Skyler Atchison',
    description: 'Cover letter',
    title: document_title,
    styles: {
      default: {
        document: {
          run: {
            size: '12pt',
            font: 'Times New Roman',
          },
          paragraph: {
            spacing: {
              after: 200,
            },
          },
        },
      },
      paragraphStyles: [
        {
          id: 'no_spacing',
          name: 'No Spacing',
          basedOn: 'default',
          paragraph: {
            spacing: {
              after: 0,
            },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: [
          new Paragraph({
            text: `Skyler Atchison`,
            style: 'no_spacing',
          }),
          new Paragraph({
            children: [
              new TextRun(`${phone} | `),
              email_link,
              new TextRun(` | ${address}`),
            ],
            style: 'no_spacing',
          }),
          new Paragraph({ text: `${date}`, style: 'no_spacing' }),
          new Paragraph({
            text: `${company_name}`,
            style: 'no_spacing',
          }),
        ],
      },
      {
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: [
          new Paragraph({
            text: 'Dear Hiring Manager,',
          }),
          new Paragraph({
            text:
              `Last fall, I created and presented BarterBuddy, an innovative cloud-backed project which was commended by my professor ` +
              `as reminiscent of a Silicon Valley startup. This experience, part of my four years of studying Computer Science, ` +
              `has equipped me with proficiency in software engineering, and I am eager to gain experience in a professional environment. ` +
              `As someone interested in software development and always seeking to learn more, ` +
              `I would love the opportunity to have a ${role} position at ${company_name}.`,
          }),
          new Paragraph(
            `During my studies, I have gained experience in both low-level and high-level languages, ` +
              `having begun my coding journey with Python before expanding my repertoire to use C++, C, and JavaScript ` +
              `in successfully implementing challenging projects such as a simulated interpreter and a file system.`
          ),
          new Paragraph(
            `One of my significant achievements was on the development team for the BarterBuddy Android app,` +
              ` which we designed and created as a Software Development project using Agile methodologies.` +
              ` This experience enhanced my skills in Java and JUnit, REST APIs, git, GitHub, database queries, feature development,` +
              ` code reviews, refactoring, and collaborative coding environments. As a Lead Computer Science Tutor,` +
              ` I further honed my communication and collaboration skills, providing a solid foundation for working effectively in a team.`
          ),
          new Paragraph(
            `I bring experience in team full-stack development, unit testing and debugging, front-end frameworks, problem solving,` +
              ` and independent learning, all of which will be essential to my role in ${role}, should you choose to hire me.` +
              ` I am confident in my ability to meaningfully contribute to the development of ${company_name}’s ${company_projects}.`
          ),
          new Paragraph({
            children: [
              new TextRun(
                `Thank you for considering me for the position. I am eager to discuss how my background aligns with the needs of ${company_name} in more detail.` +
                  ` Please feel free to contact me at ${phone} or `
              ),
              email_link,
              new TextRun(
                `. I look forward to the opportunity for an interview.`
              ),
            ],
          }),
          new Paragraph({ text: `Sincerely,`, style: 'no_spacing' }),
          new Paragraph(`Skyler Atchison`),
        ],
      },
    ],
  })

  const buffer = await Packer.toBuffer(doc)
  fs.writeFileSync(`${output_directory}/${document_title}.docx`, buffer)
}

const fill_ins = await set_options(defaults)
make_document(fill_ins)
