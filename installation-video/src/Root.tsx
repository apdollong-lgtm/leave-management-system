import { Composition, Folder, Sequence, Still } from "remotion";
import { Cover } from "./Cover";
import { Intro } from "./scenes/Intro";
import { Project } from "./scenes/Project";
import { Security } from "./scenes/Security";
import { Configuration } from "./scenes/Configuration";
import { Setup } from "./scenes/Setup";
import { Deploy } from "./scenes/Deploy";
import { Users } from "./scenes/Users";
import { TestFlow } from "./scenes/TestFlow";
import { Outro } from "./scenes/Outro";
export const InstallationGuide = () => (
  <>
    <Sequence name="Introduction" durationInFrames={240}>
      <Intro />
    </Sequence>
    <Sequence name="Create project" from={240} durationInFrames={420}>
      <Project />
    </Sequence>
    <Sequence name="Admin security" from={660} durationInFrames={420}>
      <Security />
    </Sequence>
    <Sequence name="Organization" from={1080} durationInFrames={420}>
      <Configuration />
    </Sequence>
    <Sequence name="Setup sheets" from={1500} durationInFrames={420}>
      <Setup />
    </Sequence>
    <Sequence name="Deploy web app" from={1920} durationInFrames={420}>
      <Deploy />
    </Sequence>
    <Sequence name="Create users" from={2340} durationInFrames={420}>
      <Users />
    </Sequence>
    <Sequence name="Test workflow" from={2760} durationInFrames={420}>
      <TestFlow />
    </Sequence>
    <Sequence name="Handover" from={3180} durationInFrames={240}>
      <Outro />
    </Sequence>
  </>
);
export const RemotionRoot = () => (
  <>
    <Still id="ProductCover" component={Cover} width={1080} height={1920} />
    <Composition
      id="InstallationGuide"
      component={InstallationGuide}
      durationInFrames={3420}
      fps={30}
      width={1920}
      height={1080}
    />
    <Folder name="Scenes">
      <Composition
        id="Intro"
        component={Intro}
        durationInFrames={240}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="Project"
        component={Project}
        durationInFrames={420}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="Security"
        component={Security}
        durationInFrames={420}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="Configuration"
        component={Configuration}
        durationInFrames={420}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="Setup"
        component={Setup}
        durationInFrames={420}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="Deploy"
        component={Deploy}
        durationInFrames={420}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="Users"
        component={Users}
        durationInFrames={420}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="TestFlow"
        component={TestFlow}
        durationInFrames={420}
        fps={30}
        width={1920}
        height={1080}
      />
      <Composition
        id="Outro"
        component={Outro}
        durationInFrames={240}
        fps={30}
        width={1920}
        height={1080}
      />
    </Folder>
  </>
);
